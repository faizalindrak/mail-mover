package main

import (
	"context"
	"errors"
	"fmt"
	"os"
	"os/exec"
	"path/filepath"
	goruntime "runtime"
	"sort"
	"strings"
	"sync"
	"time"

	"github.com/wailsapp/wails/v2/pkg/runtime"
)

type App struct {
	ctx       context.Context
	mu        sync.Mutex
	running   bool
	cancelled bool
}

type EnvironmentStatus struct {
	Windows        bool     `json:"windows"`
	OutlookRunning bool     `json:"outlookRunning"`
	PSTNames       []string `json:"pstNames"`
	Message        string   `json:"message"`
}

type MigrationRequest struct {
	SourcePath     string `json:"sourcePath"`
	StagingPath    string `json:"stagingPath"`
	TargetPSTName  string `json:"targetPSTName"`
	MarkAsRead     bool   `json:"markAsRead"`
	PreventSleep   bool   `json:"preventSleep"`
	SkipConversion bool   `json:"skipConversion"`
	KeepStagingEML bool   `json:"keepStagingEML"`
}

type MigrationProgress struct {
	Stage     string `json:"stage"`
	Status    string `json:"status"`
	Message   string `json:"message"`
	Folder    string `json:"folder"`
	Current   int    `json:"current"`
	Total     int    `json:"total"`
	Converted int    `json:"converted"`
	Imported  int    `json:"imported"`
	Failed    int    `json:"failed"`
}

func NewApp() *App {
	return &App{}
}

func (a *App) startup(ctx context.Context) {
	a.ctx = ctx
}

func (a *App) SelectThunderbirdFolder() (string, error) {
	return runtime.OpenDirectoryDialog(a.ctx, runtime.OpenDialogOptions{
		Title: "Choose Thunderbird Local Folders",
	})
}

func (a *App) SelectStagingFolder() (string, error) {
	return runtime.OpenDirectoryDialog(a.ctx, runtime.OpenDialogOptions{
		Title:                "Choose a temporary EML folder",
		CanCreateDirectories: true,
	})
}

func (a *App) GetEnvironmentStatus() EnvironmentStatus {
	if goruntime.GOOS != "windows" {
		return EnvironmentStatus{Message: "PST import requires Windows and Classic Outlook."}
	}
	result := runPowerShell(`
$ErrorActionPreference = 'Stop'
$outlook = [Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application')
$namespace = $outlook.GetNamespace('MAPI')
$names = @()
for ($i = 1; $i -le $namespace.Folders.Count; $i++) { $names += $namespace.Folders.Item($i).Name }
[Console]::OutputEncoding = [Text.Encoding]::UTF8
$names -join [Environment]::NewLine
`)
	if result.err != nil {
		return EnvironmentStatus{Windows: true, Message: "Open Classic Outlook, then click Check again."}
	}
	var names []string
	for _, name := range strings.Split(strings.TrimSpace(result.output), "\n") {
		name = strings.TrimSpace(name)
		if name != "" {
			names = append(names, name)
		}
	}
	sort.Strings(names)
	return EnvironmentStatus{
		Windows:        true,
		OutlookRunning: true,
		PSTNames:       names,
		Message:        "Classic Outlook is ready.",
	}
}

func (a *App) StartMigration(request MigrationRequest) error {
	if err := validateRequest(request); err != nil {
		return err
	}
	a.mu.Lock()
	if a.running {
		a.mu.Unlock()
		return errors.New("a migration is already running")
	}
	a.running = true
	a.cancelled = false
	a.mu.Unlock()

	go a.runMigration(request)
	return nil
}

func (a *App) CancelMigration() {
	a.mu.Lock()
	a.cancelled = true
	a.mu.Unlock()
}

func (a *App) OpenFolder(path string) error {
	path = filepath.Clean(strings.TrimSpace(path))
	if path == "." || path == "" {
		return errors.New("choose a folder first")
	}
	if goruntime.GOOS == "windows" {
		return newBackgroundCommand("explorer.exe", path).Start()
	}
	if goruntime.GOOS == "darwin" {
		return exec.Command("open", path).Start()
	}
	return exec.Command("xdg-open", path).Start()
}

func (a *App) runMigration(request MigrationRequest) {
	defer func() {
		a.mu.Lock()
		a.running = false
		a.mu.Unlock()
	}()

	progress := MigrationProgress{Stage: "prepare", Status: "running", Message: "Checking the selected folders."}
	a.emit(progress)

	if request.PreventSleep {
		preventSleep(true)
		defer preventSleep(false)
	}

	if !request.SkipConversion {
		if err := prepareStagingFolder(request.StagingPath); err != nil {
			a.fail(progress, err)
			return
		}
		progress.Stage = "convert"
		progress.Message = "Converting Thunderbird mailboxes to EML."
		a.emit(progress)
		converted, err := convertMboxTree(request.SourcePath, request.StagingPath, func(folder string, count int) {
			progress.Folder = folder
			progress.Converted += count
			progress.Message = fmt.Sprintf("Converted %s", folder)
			a.emit(progress)
		})
		progress.Converted = converted
		if err != nil {
			a.fail(progress, err)
			return
		}
	}

	if a.isCancelled() {
		progress.Status = "cancelled"
		progress.Message = "Migration cancelled before Outlook import."
		a.emit(progress)
		return
	}

	files, err := collectEMLFiles(request.StagingPath)
	if err != nil {
		a.fail(progress, err)
		return
	}
	if len(files) == 0 {
		a.fail(progress, errors.New("no EML files were found to import"))
		return
	}

	progress.Stage = "import"
	progress.Total = len(files)
	progress.Message = "Importing messages into Classic Outlook."
	a.emit(progress)

	importerPath, err := bundledImporterPath()
	if err != nil {
		a.fail(progress, fmt.Errorf("prepare Outlook importer: %w", err))
		return
	}
	cancelFile := filepath.Join(os.TempDir(), fmt.Sprintf("mbox2pst-cancel-%d", time.Now().UnixNano()))
	args := []string{
		"-NoProfile", "-ExecutionPolicy", "Bypass", "-File", importerPath,
		"-SourcePath", request.StagingPath,
		"-TargetPSTName", request.TargetPSTName,
		"-CancelFile", cancelFile,
	}
	if request.MarkAsRead {
		args = append(args, "-MarkAsRead")
	}
	command := newBackgroundCommand("powershell.exe", args...)
	stdout, err := command.StdoutPipe()
	if err != nil {
		a.fail(progress, err)
		return
	}
	command.Stderr = command.Stdout
	if err := command.Start(); err != nil {
		a.fail(progress, fmt.Errorf("start Outlook importer: %w", err))
		return
	}

	lines := make(chan string)
	go scanLines(stdout, lines)
	for line := range lines {
		if a.isCancelled() {
			_ = os.WriteFile(cancelFile, []byte("cancel"), 0o600)
		}
		parseImporterLine(line, &progress)
		a.emit(progress)
	}
	_ = os.Remove(cancelFile)
	if err := command.Wait(); err != nil && progress.Status != "cancelled" {
		a.fail(progress, fmt.Errorf("Outlook import stopped: %w", err))
		return
	}
	if a.isCancelled() && progress.Status == "running" {
		progress.Status = "cancelled"
	}
	if progress.Status == "running" && progress.Imported+progress.Failed < progress.Total {
		a.fail(progress, errors.New("Outlook import ended before every EML file was processed"))
		return
	}
	if progress.Status == "cancelled" {
		progress.Message = "Migration cancelled. Already imported messages remain in Outlook."
		a.emit(progress)
		return
	}

	if !request.KeepStagingEML && !request.SkipConversion {
		progress.Stage = "cleanup"
		progress.Message = "Removing temporary EML files."
		a.emit(progress)
		if err := os.RemoveAll(request.StagingPath); err != nil {
			progress.Message = "Migration finished, but the temporary EML folder could not be removed."
		}
	}

	progress.Stage = "complete"
	progress.Status = "complete"
	progress.Message = fmt.Sprintf("Import complete: %d messages imported into %s.", progress.Imported, request.TargetPSTName)
	a.emit(progress)
}

func validateRequest(request MigrationRequest) error {
	if goruntime.GOOS != "windows" {
		return errors.New("PST import requires Windows and Classic Outlook")
	}
	if strings.TrimSpace(request.StagingPath) == "" {
		return errors.New("choose a temporary EML folder")
	}
	if !request.SkipConversion && strings.TrimSpace(request.SourcePath) == "" {
		return errors.New("choose your Thunderbird Local Folders folder")
	}
	if strings.TrimSpace(request.TargetPSTName) == "" {
		return errors.New("choose an Outlook PST")
	}
	if !request.SkipConversion {
		source := filepath.Clean(request.SourcePath)
		staging := filepath.Clean(request.StagingPath)
		if source == staging || pathContains(source, staging) {
			return errors.New("choose a temporary EML folder outside the Thunderbird source folder")
		}
	} else {
		info, err := os.Stat(request.StagingPath)
		if err != nil || !info.IsDir() {
			return errors.New("choose an existing folder that contains EML files")
		}
	}
	return nil
}

func (a *App) emit(progress MigrationProgress) {
	runtime.EventsEmit(a.ctx, "migration:progress", progress)
}

func (a *App) fail(progress MigrationProgress, err error) {
	progress.Status = "error"
	progress.Message = err.Error()
	a.emit(progress)
}

func (a *App) isCancelled() bool {
	a.mu.Lock()
	defer a.mu.Unlock()
	return a.cancelled
}

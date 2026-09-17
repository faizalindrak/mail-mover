package main

import (
	"bufio"
	"errors"
	"fmt"
	"os"
	"os/exec"
	"path/filepath"
	"strconv"
	"strings"
)

type commandResult struct {
	output string
	err    error
}

func runPowerShell(script string) commandResult {
	command := exec.Command("powershell.exe", "-NoProfile", "-NonInteractive", "-Command", script)
	output, err := command.CombinedOutput()
	return commandResult{output: string(output), err: err}
}

func prepareStagingFolder(path string) error {
	info, err := os.Stat(path)
	if errors.Is(err, os.ErrNotExist) {
		return os.MkdirAll(path, 0o755)
	}
	if err != nil {
		return fmt.Errorf("open temporary EML folder: %w", err)
	}
	if !info.IsDir() {
		return errors.New("the temporary EML path must be a folder")
	}
	entries, err := os.ReadDir(path)
	if err != nil {
		return fmt.Errorf("read temporary EML folder: %w", err)
	}
	if len(entries) != 0 {
		return errors.New("choose an empty temporary EML folder to prevent mixing this migration with existing files")
	}
	return nil
}

func collectEMLFiles(source string) ([]string, error) {
	var files []string
	err := filepath.WalkDir(source, func(path string, entry os.DirEntry, walkErr error) error {
		if walkErr != nil {
			return walkErr
		}
		if !entry.IsDir() && strings.EqualFold(filepath.Ext(entry.Name()), ".eml") {
			files = append(files, path)
		}
		return nil
	})
	return files, err
}

func bundledImporterPath() (string, error) {
	directory := filepath.Join(os.TempDir(), "MailboxMover")
	if err := os.MkdirAll(directory, 0o700); err != nil {
		return "", err
	}
	path := filepath.Join(directory, "Import-ThunderbirdToOutlook.ps1")
	if err := os.WriteFile(path, importerScript, 0o600); err != nil {
		return "", err
	}
	return path, nil
}

func scanLines(reader interface{ Read([]byte) (int, error) }, output chan<- string) {
	defer close(output)
	scanner := bufio.NewScanner(reader)
	scanner.Buffer(make([]byte, 64*1024), 1024*1024)
	for scanner.Scan() {
		output <- scanner.Text()
	}
}

func parseImporterLine(line string, progress *MigrationProgress) {
	parts := strings.Split(line, "|")
	if len(parts) == 0 {
		return
	}
	switch parts[0] {
	case "PROGRESS":
		if len(parts) < 5 {
			return
		}
		progress.Current, _ = strconv.Atoi(parts[1])
		progress.Imported, _ = strconv.Atoi(parts[2])
		progress.Failed, _ = strconv.Atoi(parts[3])
		progress.Folder = parts[4]
		progress.Message = fmt.Sprintf("Importing %s (%d of %d)", progress.Folder, progress.Current, progress.Total)
	case "FAILED":
		progress.Failed++
		if len(parts) > 1 {
			progress.Message = "Skipped a message: " + parts[1]
		}
	case "CANCELLED":
		progress.Status = "cancelled"
	case "COMPLETE":
		if len(parts) >= 3 {
			progress.Imported, _ = strconv.Atoi(parts[1])
			progress.Failed, _ = strconv.Atoi(parts[2])
		}
	case "STATUS":
		if len(parts) > 1 {
			progress.Message = parts[1]
		}
	}
}

package main

import (
	"errors"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"sort"
	"strings"

	mbox "github.com/emersion/go-mbox"
)

type folderProgress func(folder string, converted int)

// convertMboxTree extracts every Thunderbird mbox file under source into EML
// files under destination while preserving the mailbox hierarchy.
func convertMboxTree(source, destination string, progress folderProgress) (int, error) {
	source = filepath.Clean(source)
	destination = filepath.Clean(destination)
	if source == destination || pathContains(source, destination) {
		return 0, errors.New("choose an EML output folder outside the Thunderbird source folder")
	}

	files, err := discoverMboxFiles(source)
	if err != nil {
		return 0, err
	}
	if len(files) == 0 {
		return 0, errors.New("no Thunderbird mbox files were found in the selected folder")
	}

	total := 0
	for _, path := range files {
		relative, err := filepath.Rel(source, path)
		if err != nil {
			return total, fmt.Errorf("resolve mailbox path %q: %w", path, err)
		}
		outputDir := filepath.Join(destination, normaliseThunderbirdPath(relative))
		converted, err := convertMboxFile(path, outputDir)
		if err != nil {
			return total, fmt.Errorf("convert %q: %w", relative, err)
		}
		total += converted
		if progress != nil {
			progress(relative, converted)
		}
	}
	return total, nil
}

func discoverMboxFiles(source string) ([]string, error) {
	info, err := os.Stat(source)
	if err != nil {
		return nil, fmt.Errorf("open Thunderbird source: %w", err)
	}
	if !info.IsDir() {
		return nil, errors.New("the Thunderbird source must be a folder")
	}

	var files []string
	err = filepath.WalkDir(source, func(path string, entry os.DirEntry, walkErr error) error {
		if walkErr != nil {
			return walkErr
		}
		if entry.IsDir() {
			return nil
		}
		name := strings.ToLower(entry.Name())
		if strings.HasSuffix(name, ".msf") || strings.HasSuffix(name, ".eml") || strings.HasSuffix(name, ".sqlite") || strings.HasSuffix(name, ".json") || strings.HasSuffix(name, ".dat") || strings.HasSuffix(name, ".html") {
			return nil
		}
		file, err := os.Open(path)
		if err != nil {
			return err
		}
		buffer := make([]byte, 5)
		_, readErr := io.ReadFull(file, buffer)
		_ = file.Close()
		if readErr != nil && !errors.Is(readErr, io.ErrUnexpectedEOF) {
			return readErr
		}
		if string(buffer) == "From " {
			files = append(files, path)
		}
		return nil
	})
	sort.Strings(files)
	return files, err
}

func convertMboxFile(path, outputDir string) (int, error) {
	input, err := os.Open(path)
	if err != nil {
		return 0, err
	}
	defer input.Close()

	if err := os.MkdirAll(outputDir, 0o755); err != nil {
		return 0, err
	}

	reader := mbox.NewReader(input)
	converted := 0
	for {
		message, err := reader.NextMessage()
		if errors.Is(err, io.EOF) {
			break
		}
		if err != nil {
			return converted, err
		}
		outputPath := filepath.Join(outputDir, fmt.Sprintf("%06d.eml", converted+1))
		output, err := os.Create(outputPath)
		if err != nil {
			return converted, err
		}
		_, copyErr := io.Copy(output, message)
		closeErr := output.Close()
		if copyErr != nil {
			return converted, copyErr
		}
		if closeErr != nil {
			return converted, closeErr
		}
		converted++
	}
	return converted, nil
}

func normaliseThunderbirdPath(relative string) string {
	parts := strings.Split(filepath.Clean(relative), string(filepath.Separator))
	for index, part := range parts {
		parts[index] = strings.TrimSuffix(part, ".sbd")
	}
	return filepath.Join(parts...)
}

func pathContains(parent, candidate string) bool {
	relative, err := filepath.Rel(parent, candidate)
	return err == nil && relative != "." && relative != ".." && !strings.HasPrefix(relative, ".."+string(filepath.Separator))
}

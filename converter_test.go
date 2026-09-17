package main

import (
	"os"
	"path/filepath"
	"testing"
)

const sampleMbox = `From first@example.com Thu Jan  1 00:00:01 2015
From: first@example.com
Subject: First

First body.

From second@example.com Thu Jan  1 00:00:02 2015
From: second@example.com
Subject: Second

Second body.
`

func TestConvertMboxTreePreservesThunderbirdFolders(t *testing.T) {
	source := t.TempDir()
	destination := t.TempDir()
	mailboxDir := filepath.Join(source, "Inbox.sbd")
	if err := os.MkdirAll(mailboxDir, 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(mailboxDir, "Work"), []byte(sampleMbox), 0o600); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(mailboxDir, "Work.msf"), []byte("index"), 0o600); err != nil {
		t.Fatal(err)
	}

	converted, err := convertMboxTree(source, destination, nil)
	if err != nil {
		t.Fatal(err)
	}
	if converted != 2 {
		t.Fatalf("converted = %d, want 2", converted)
	}
	for _, filename := range []string{"000001.eml", "000002.eml"} {
		if _, err := os.Stat(filepath.Join(destination, "Inbox", "Work", filename)); err != nil {
			t.Errorf("missing %s: %v", filename, err)
		}
	}
}

func TestConvertMboxTreeRejectsDestinationInsideSource(t *testing.T) {
	source := t.TempDir()
	_, err := convertMboxTree(source, filepath.Join(source, "output"), nil)
	if err == nil {
		t.Fatal("expected destination validation error")
	}
}

func TestPrepareStagingFolderRejectsExistingContents(t *testing.T) {
	folder := t.TempDir()
	if err := os.WriteFile(filepath.Join(folder, "existing.txt"), []byte("keep me"), 0o600); err != nil {
		t.Fatal(err)
	}
	if err := prepareStagingFolder(folder); err == nil {
		t.Fatal("expected non-empty staging folder error")
	}
}

func TestDiscoverMboxFilesIgnoresMetadata(t *testing.T) {
	source := t.TempDir()
	if err := os.WriteFile(filepath.Join(source, "Inbox"), []byte(sampleMbox), 0o600); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(source, "Inbox.msf"), []byte(sampleMbox), 0o600); err != nil {
		t.Fatal(err)
	}

	files, err := discoverMboxFiles(source)
	if err != nil {
		t.Fatal(err)
	}
	if len(files) != 1 || filepath.Base(files[0]) != "Inbox" {
		t.Fatalf("files = %v, want Inbox only", files)
	}
}

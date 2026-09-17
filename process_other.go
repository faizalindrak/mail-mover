//go:build !windows

package main

import "os/exec"

func configureBackgroundCommand(*exec.Cmd) {}

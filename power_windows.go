//go:build windows

package main

import "golang.org/x/sys/windows"

const (
	esContinuous      = 0x80000000
	esSystemRequired  = 0x00000001
	esDisplayRequired = 0x00000002
)

var setThreadExecutionState = windows.NewLazySystemDLL("kernel32.dll").NewProc("SetThreadExecutionState")

func preventSleep(enabled bool) {
	state := uintptr(esContinuous)
	if enabled {
		state |= esSystemRequired | esDisplayRequired
	}
	_, _, _ = setThreadExecutionState.Call(state)
}

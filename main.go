package main

import (
	"embed"

	"github.com/wailsapp/wails/v2"
	"github.com/wailsapp/wails/v2/pkg/options"
	"github.com/wailsapp/wails/v2/pkg/options/assetserver"
	"github.com/wailsapp/wails/v2/pkg/options/windows"
)

//go:embed all:frontend/dist
var assets embed.FS

//go:embed scripts/Import-ThunderbirdToOutlook.ps1
var importerScript []byte

func main() {
	app := NewApp()

	err := wails.Run(&options.App{
		Title:     "Mailbox Mover",
		Width:     1180,
		Height:    820,
		MinWidth:  900,
		MinHeight: 680,
		AssetServer: &assetserver.Options{
			Assets: assets,
		},
		BackgroundColour: options.NewRGB(246, 247, 249),
		OnStartup:        app.startup,
		Bind: []interface{}{
			app,
		},
		Windows: &windows.Options{
			Theme:        windows.SystemDefault,
			BackdropType: windows.Mica,
		},
	})
	if err != nil {
		println("Error:", err.Error())
	}
}

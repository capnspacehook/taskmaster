// +build windows

package main

import (
	"github.com/capnspacehook/taskmaster"
)

func main() {
	var err error

	taskService, err := taskmaster.Connect("", "", "", "")
	if err != nil {
		panic(err)
	}
	defer taskService.Disconnect()

	_, err = taskService.DeleteFolder("\\NewFolder", true)
	if err != nil {
		panic(err)
	}
}

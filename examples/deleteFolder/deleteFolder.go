// +build windows

package main

import (
	"github.com/capnspacehook/taskmaster"
)

func main() {
	var err error
	var taskService taskmaster.TaskService

	err = taskService.Connect("", "", "", "")
	if err != nil {
		panic(err)
	}
	defer taskService.Cleanup()

	_, err = taskService.DeleteFolder("\\NewFolder", true)
	if err != nil {
		panic(err)
	}
}

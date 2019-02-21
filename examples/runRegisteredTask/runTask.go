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
	defer taskService.Cleanup()

	task, err := taskService.GetRegisteredTask("\\NewTask")
	if err != nil {
		panic(err)
	}

	if task != nil {
		runningTask, err := task.Run(nil)
		if err != nil {
			panic(err)
		}
		defer runningTask.Release()
	}
}

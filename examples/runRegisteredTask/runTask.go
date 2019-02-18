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

	task, err := taskService.GetRegisteredTask("\\NewTask")
	if err != nil {
		panic(err)
	}

	if task != nil {
		defer task.Release()
		task.Run([]string{"/c", "timeout 69"}, taskmaster.TASK_RUN_AS_SELF, 0, "")
	}
}

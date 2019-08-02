// +build windows

package main

import (
	"fmt"

	"github.com/capnspacehook/taskmaster"
)

func main() {
	var err error

	taskService, err := taskmaster.Connect("", "", "", "")
	if err != nil {
		panic(err)
	}
	defer taskService.Disconnect()

	tasks, err := taskService.GetRegisteredTasks()
	if err != nil {
		panic(err)
	}

	fmt.Println("ABUSABLE TASKS:")
	for _, task := range tasks {
		for _, action := range task.Definition.Actions {
			if action.GetType() == taskmaster.TASK_ACTION_COM_HANDLER {
				for _, trigger := range task.Definition.Triggers {
					if trigger.GetType() == taskmaster.TASK_TRIGGER_LOGON {
						fmt.Printf("Name: %s\n", task.Name)
						fmt.Printf("Path: %s\n", task.Path)
						fmt.Printf("Context: %s\n", task.Definition.Context)
						fmt.Printf("CLSID: %s\n\n", action.(taskmaster.ComHandlerAction).ClassID)
					}
				}
			}
		}
	}
}

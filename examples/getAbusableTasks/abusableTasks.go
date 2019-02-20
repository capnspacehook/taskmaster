// +build windows

package main

import (
	"fmt"

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

	err = taskService.GetRegisteredTasks()
	if err != nil {
		panic(err)
	}

	fmt.Println("ABUSABLE TASKS:")
	for _, task := range taskService.RegisteredTasks {
		for _, action := range task.Definition.Actions {
			if action.GetType() == taskmaster.TASK_ACTION_COM_HANDLER {
				for _, trigger := range task.Definition.Triggers {
					if trigger.GetType() == taskmaster.TASK_TRIGGER_LOGON {
						fmt.Println("\n--------------------------------------------------------")
						fmt.Printf("Name: %s\n", task.Name)
						fmt.Printf("Path: %s\n", task.Path)
						fmt.Printf("Context: %s\n", task.Definition.Context)
						fmt.Printf("CLSID: %s\n", action.(taskmaster.ComHandlerAction).ClassID)
					}
				}
			}
		}
	}
}

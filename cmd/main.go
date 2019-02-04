package main

import (
	"fmt"

	"github.com/capnspacehook/taskmaster"
)

func main() {
	var err error
	var taskService taskmaster.TaskService
	err = taskService.Connect()
	if err != nil {
		panic(err)
	}
	defer taskService.Disconnect()

	tasks, err := taskService.GetRegisteredTasks()
	if err != nil {
		panic(err)
	}

	for _, task := range(tasks) {
		fmt.Println("----------------------------------------------")
		fmt.Printf("Name: %s\n", task.Name)
		fmt.Printf("Path: %s\n", task.Path)
		fmt.Printf("Context: %s\n", task.Definition.Context)

		fmt.Println("Principle:")
		fmt.Printf("\tName: %s\n", task.Definition.Principal.Name)
		fmt.Printf("\tGroupID: %s\n", task.Definition.Principal.GroupID)
		fmt.Printf("\tID: %s\n", task.Definition.Principal.ID)
		fmt.Printf("\tLogon Type: %d\n", task.Definition.Principal.LogonType)
		fmt.Printf("\tRunLevel: %d\n", task.Definition.Principal.RunLevel)
		fmt.Printf("\tUserID: %s\n", task.Definition.Principal.UserID)

		for i, action := range(task.Definition.Actions) {
			fmt.Printf("Action %d:\n", i + 1)
			switch action.GetType() {
			case taskmaster.TASK_ACTION_EXEC:
				execAction := action.(taskmaster.ExecAction)
				fmt.Printf("\tPath: %s\n", execAction.Path)
				fmt.Printf("\tArgs: %s\n", execAction.Args)
			case taskmaster.TASK_ACTION_COM_HANDLER, taskmaster.TASK_ACTION_CUSTOM_HANDLER:
				comHandlerAction := action.(taskmaster.ComHandlerAction)
				fmt.Printf("\tClassID: %s\n", comHandlerAction.ClassID)
				fmt.Printf("\tData: %s\n", comHandlerAction.Data)
			}
		}

		fmt.Printf("Enabled: %t\n", task.Enabled)
		fmt.Printf("Number of Missed Runs: %d\n", task.MissedRuns)
		fmt.Printf("Next Run Time: %s\n", task.NextRunTime)
		fmt.Printf("Last Run Time: %s\n", task.LastRunTime)
		fmt.Printf("Last Task Result %d\n", task.LastTaskResult)
	}
}

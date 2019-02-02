package main

import (
	"fmt"
	//"reflect"

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
		fmt.Printf("Name: %s\n", task.Name)
		fmt.Printf("Path: %s\n", task.Path)
		fmt.Printf("Context: %s\n", task.Definition.Actions.Context)

		/*for _, action := range(task.Definition.Actions.Actions) {
			switch reflect.ValueOf(&action).Elem().FieldByName("Type").Int() {
			case taskmaster.TASK_ACTION_EXEC:
				execAction := action.(taskmaster.ExecAction)
				fmt.Printf("Path: %s\n", execAction.Path)
				fmt.Printf("Args: %s\n", execAction.Args)
			case taskmaster.TASK_ACTION_COM_HANDLER:
				comHandlerAction := action.(taskmaster.ComHandlerAction)
				fmt.Printf("ClassID: %s\n", comHandlerAction.ClassID)
				fmt.Printf("Data: %s\n", comHandlerAction.Data)
			}

		}*/

		fmt.Printf("Enabled: %t\n", task.Enabled)
		fmt.Printf("Number of Missed Runs: %d\n", task.MissedRuns)
		fmt.Printf("Next Run Time: %s\n", task.NextRunTime)
		fmt.Printf("Last Run Time: %s\n", task.LastRunTime)
		fmt.Printf("Last Task Result %d\n", task.LastTaskResult)
	}
}

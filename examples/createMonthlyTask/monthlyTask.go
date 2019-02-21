// +build windows

package main

import (
	"time"

	"github.com/capnspacehook/taskmaster"
)

func main() {
	var err error

	taskService, err := taskmaster.Connect("", "", "", "")
	if err != nil {
		panic(err)
	}
	defer taskService.Cleanup()

	newTaskDef := taskService.NewTaskDefinition()

	newTaskDef.AddExecAction("calc.exe", "", "", "")
	newTaskDef.AddMonthlyTrigger(17, taskmaster.February, "", time.Now().Add(5*time.Second))
	newTaskDef.RegistrationInfo.Author = "capnspacehook"
	newTaskDef.RegistrationInfo.Description = "Pops calc. What else would you expect?"

	_, _, err = taskService.CreateTask("\\NewFolder\\NewTask", newTaskDef, true)
	if err != nil {
		panic(err)
	}
}

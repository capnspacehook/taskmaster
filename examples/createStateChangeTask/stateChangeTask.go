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

	newTaskDef := taskService.NewTaskDefinition()

	newTaskDef.AddExecAction("cmd.exe", "/c echo Hi l33t h4x0r; timeout 5", "", "Launch CMD")
	newTaskDef.AddSessionStateChangeTrigger("", taskmaster.TASK_SESSION_LOCK, "PT5S")
	newTaskDef.Principal.RunLevel = taskmaster.TASK_RUNLEVEL_HIGHEST
	newTaskDef.RegistrationInfo.Author = "capnspacehook"
	newTaskDef.RegistrationInfo.Description = "CMD greets you when you logon :)"

	_, _, err = taskService.CreateTask("\\NewFolder\\Greeter", newTaskDef, true)
	if err != nil {
		panic(err)
	}
}

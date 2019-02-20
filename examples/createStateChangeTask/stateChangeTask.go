package main

import (
	"time"

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

	newTaskDef := taskService.NewTaskDefinition()

	newTaskDef.AddExecAction("cmd.exe", "/c echo Hi l33t h4x0r; timeout 5", "", "Launch CMD")
	newTaskDef.AddSessionStateChangeTrigger("", taskmaster.TASK_SESSION_LOCK, "PT5S", "", time.Time{}, time.Time{}, "", "", "", false, true)
	newTaskDef.Principal.RunLevel = taskmaster.TASK_RUNLEVEL_HIGHEST
	newTaskDef.RegistrationInfo.Author = "capnspacehook"
	newTaskDef.RegistrationInfo.Description = "CMD greets you when you logon :)"

	_, _, err = taskService.CreateTask("\\NewFolder\\Greeter", newTaskDef, "", "", newTaskDef.Principal.LogonType, true)
	if err != nil {
		panic(err)
	}
}

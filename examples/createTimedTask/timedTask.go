package main

import (
	"fmt"
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

	newTaskDef.AddExecAction("cmd.exe", "/c $(Arg0)", "", "Launch CMD")
	newTaskDef.AddExecAction("calc.exe", "", "", "Pop Calc")
	newTaskDef.AddTimeTrigger("", "Run in 5 Seconds", time.Now().Add(5*time.Second), time.Time{}, "", "", "", false, true)
	newTaskDef.RegistrationInfo.Author = "capnspacehook"
	newTaskDef.RegistrationInfo.Description = "Double trouble... cmd.exe and calc. l33t h4x0rs must be at it again..."

	_, err = taskService.CreateTask("\\NewFolder\\NewTask", newTaskDef, "", "", newTaskDef.Principal.LogonType, true)
	if err != nil {
		panic(err)
	}
	runningTask, err := taskService.RegisteredTasks["\\NewFolder\\NewTask"].Run([]string{"timeout 42"}, taskmaster.TASK_RUN_AS_SELF, 0, "")
	if err != nil {
		panic(err)
	}

	fmt.Printf("Running task PID: %d\n", runningTask.EnginePID)
}

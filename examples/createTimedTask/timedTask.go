// +build windows

package main

import (
	"fmt"
	"time"

	"github.com/capnspacehook/taskmaster"
	"github.com/rickb777/date/period"
)

func main() {
	var err error

	taskService, err := taskmaster.Connect("", "", "", "")
	if err != nil {
		panic(err)
	}
	defer taskService.Disconnect()

	newTaskDef := taskService.NewTaskDefinition()

	newTaskDef.AddExecAction("cmd.exe", "/c $(Arg0)", "", "Launch CMD")
	newTaskDef.AddExecAction("calc.exe", "", "", "Pop Calc")
	newTaskDef.AddTimeTrigger(period.NewHMS(0, 0, 0), time.Now().Add(5*time.Second))
	newTaskDef.RegistrationInfo.Author = "capnspacehook"
	newTaskDef.RegistrationInfo.Description = "Double trouble... cmd.exe and calc. l33t h4x0rs must be at it again..."
	newTaskDef.RegistrationInfo.Date = time.Time{}

	newTask, _, err := taskService.CreateTask("\\NewFolder\\NewTask", newTaskDef, true)
	if err != nil {
		panic(err)
	}
	runningTask, err := newTask.Run([]string{"timeout 42"})
	if err != nil {
		panic(err)
	}
	defer runningTask.Release()

	fmt.Printf("Running task PID: %d\n", runningTask.EnginePID)
}

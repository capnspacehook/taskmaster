// +build windows

package main

import (
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
	defer taskService.Cleanup()

	newTaskDef := taskService.NewTaskDefinition()

	newTaskDef.AddExecAction("calc.exe", "", "", "Pop Calc")

	// define default values for attributes of the trigger we aren't intrested in,
	// and define a time period of 5 minutes
	var (
		defaultTime   time.Time
		defaultPeriod period.Period
		repFive       = period.NewHMS(0, 5, 0)
	)

	// add a trigger that starts 5 seconds from when the task is created, and restarts calc.exe every 5 minutes
	newTaskDef.AddTimeTriggerEx(defaultPeriod, "", time.Now().Add(3*time.Second), defaultTime, defaultPeriod, defaultPeriod, repFive, true, true)

	_, _, err = taskService.CreateTask("\\NewFolder\\NewTask", newTaskDef, true)
	if err != nil {
		panic(err)
	}
}

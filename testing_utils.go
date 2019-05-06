package taskmaster

import (
	"time"

	"github.com/rickb777/date/period"
)

func createTestTask(taskSvc *TaskService) *RegisteredTask {
	newTaskDef := taskSvc.NewTaskDefinition()
	newTaskDef.AddExecAction("cmd.exe", "/c timeout $(Arg0)", "", "")
	newTaskDef.AddTimeTrigger(period.NewHMS(0, 0, 0), time.Now().Add(10*time.Hour))
	newTaskDef.Settings.MultipleInstances = TASK_INSTANCES_PARALLEL

	task, _, err := taskSvc.CreateTask("\\Taskmaster\\TestTask", newTaskDef, false)
	if err != nil {
		panic(err)
	}

	return task
}

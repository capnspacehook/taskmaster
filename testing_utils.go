// +build windows

package taskmaster

import (
	"time"
)

func createTestTask(taskSvc *TaskService) RegisteredTask {
	newTaskDef := taskSvc.NewTaskDefinition()
	newTaskDef.AddAction(ExecAction{
		Path: "cmd.exe",
		Args: "/c timeout $(Arg0)",
	})
	newTaskDef.AddTrigger(TimeTrigger{
		TaskTrigger: TaskTrigger{
			StartBoundary: time.Now().Add(10 * time.Hour),
		},
	})
	newTaskDef.Settings.MultipleInstances = TASK_INSTANCES_PARALLEL

	task, _, err := taskSvc.CreateTask("\\Taskmaster\\TestTask", newTaskDef, false)
	if err != nil {
		panic(err)
	}

	return task
}

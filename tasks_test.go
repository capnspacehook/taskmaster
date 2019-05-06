// +build windows

package taskmaster

import (
	"testing"
	"time"

	"github.com/rickb777/date/period"
)

var testTask *RegisteredTask

// create a task that can be used to test running, stopping etc
func init() {
	taskService, err := Connect("", "", "", "")
	if err != nil {
		panic(err)
	}

	newTaskDef := taskService.NewTaskDefinition()
	newTaskDef.AddExecAction("cmd.exe", "/c timeout $(Arg0)", "", "")
	newTaskDef.AddTimeTrigger(period.NewHMS(0, 0, 0), time.Now().Add(10*time.Hour))
	newTaskDef.Settings.MultipleInstances = TASK_INSTANCES_PARALLEL

	testTask, _, err = taskService.CreateTask("\\Taskmaster\\TestTask", newTaskDef, true)
	if err != nil {
		panic(err)
	}
}

func TestRunRegisteredTask(t *testing.T) {
	runningTask, err := testTask.Run([]string{"0"})
	if err != nil {
		t.Error(err)
	}
	time.Sleep(500 * time.Millisecond)
	runningTask.Release()
}

func TestRefreshRunningTask(t *testing.T) {
	runningTask, err := testTask.Run([]string{"3"})
	if err != nil {
		t.Error(err)
	}

	err = runningTask.Refresh()
	if err != nil {
		t.Error(err)
	}

	// make sure above running task is stopped
	time.Sleep(3 * time.Second)
	runningTask.Release()
}

func TestStopRunningTask(t *testing.T) {
	runningTask, err := testTask.Run([]string{"9001"})
	if err != nil {
		t.Error(err)
	}

	err = runningTask.Stop()
	if err != nil {
		t.Error(err)
	}
}

func TestGetInstancesRegisteredTask(t *testing.T) {
	var err error

	// create a few running tasks so that there will be multiple instances
	// of the registered task running
	runningTasks := make([]*RunningTask, 5, 5)
	for i := range runningTasks {
		runningTasks[i], err = testTask.Run([]string{"3"})
		if err != nil {
			t.Error(err)
		}
	}

	instances, err := testTask.GetInstances()
	if err != nil {
		t.Error(err)
	}

	if len(instances) != 5 {
		t.Errorf("should have 5 instances, got %d instead", len(instances))
	}

	for _, instance := range instances {
		if instance == nil {
			t.Error("no instances should be nil")
		}
	}

	time.Sleep(3 * time.Second)
	for _, rTask := range runningTasks {
		rTask.Release()
	}
}

func TestStopRegisteredTask(t *testing.T) {
	var err error

	for i := 0; i < 5; i++ {
		_, err = testTask.Run([]string{"3"})
		if err != nil {
			t.Error(err)
		}
	}

	allStopped := testTask.Stop()
	if !allStopped {
		t.Error("all tasks should have stopped")
	}
}

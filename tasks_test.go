// +build windows

package taskmaster

import (
	"testing"
	"time"
)

func TestRelease(t *testing.T) {
	var rt RunningTask
	rt.Release()
}

func TestRunRegisteredTask(t *testing.T) {
	taskService, err := Connect()
	if err != nil {
		t.Error(err)
	}
	testTask := createTestTask(taskService)
	defer taskService.Disconnect()

	runningTask, err := testTask.Run([]string{"0"})
	if err != nil {
		t.Error(err)
	}
	time.Sleep(500 * time.Millisecond)
	runningTask.Release()
}

func TestRefreshRunningTask(t *testing.T) {
	taskService, err := Connect()
	if err != nil {
		t.Error(err)
	}
	testTask := createTestTask(taskService)
	defer taskService.Disconnect()

	runningTask, err := testTask.Run([]string{"1"})
	if err != nil {
		t.Error(err)
	}

	err = runningTask.Refresh()
	if err != nil {
		t.Error(err)
	}

	// make sure above running task is stopped
	time.Sleep(time.Second)
	runningTask.Release()
}

func TestStopRunningTask(t *testing.T) {
	taskService, err := Connect()
	if err != nil {
		t.Error(err)
	}
	testTask := createTestTask(taskService)
	defer taskService.Disconnect()

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
	taskService, err := Connect()
	if err != nil {
		t.Error(err)
	}
	testTask := createTestTask(taskService)
	defer taskService.Disconnect()

	runningTasks := make(RunningTaskCollection, 5, 5)

	// create a few running tasks so that there will be multiple instances
	// of the registered task running
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

	time.Sleep(3 * time.Second)
	runningTasks.Release()
	instances.Release()
}

func TestStopRegisteredTask(t *testing.T) {
	taskService, err := Connect()
	if err != nil {
		t.Error(err)
	}
	testTask := createTestTask(taskService)
	defer taskService.Disconnect()

	for i := 0; i < 5; i++ {
		_, err := testTask.Run([]string{"3"})
		if err != nil {
			t.Error(err)
		}
	}

	err = testTask.Stop()
	if err != nil {
		t.Errorf("error stopping tasks: %v", err)
	}
}

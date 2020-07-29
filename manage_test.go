// +build windows

package taskmaster

import (
	"strings"
	"testing"
	"time"

	"github.com/rickb777/date/period"
)

func TestLocalConnect(t *testing.T) {
	taskService, err := Connect()
	if err != nil {
		t.Error(err)
	}
	taskService.Disconnect()
}

func TestCreateTask(t *testing.T) {
	var err error
	taskService, err := Connect()
	if err != nil {
		t.Error(err)
	}
	defer taskService.Disconnect()

	// test ExecAction
	execTaskDef := taskService.NewTaskDefinition()
	execTaskDef.AddExecAction("calc.exe", "", "", "")

	_, _, err = taskService.CreateTask("\\Taskmaster\\ExecAction", execTaskDef, true)
	if err != nil {
		t.Error(err)
	}

	defaultPeriod := period.NewHMS(0, 5, 0)

	// test ComHandlerAction
	comHandlerDef := taskService.NewTaskDefinition()
	comHandlerDef.AddComHandlerAction("{F0001111-0000-0000-0000-0000FEEDACDC}", "", "")

	_, _, err = taskService.CreateTask("\\Taskmaster\\ComHandlerAction", comHandlerDef, true)
	if err != nil {
		t.Error(err)
	}

	// test BootTrigger
	bootTriggerDef := taskService.NewTaskDefinition()
	bootTriggerDef.AddExecAction("calc.exe", "", "", "")
	bootTriggerDef.AddBootTrigger(defaultPeriod)
	_, _, err = taskService.CreateTask("\\Taskmaster\\BootTrigger", bootTriggerDef, true)
	if err != nil {
		t.Error(err)
	}

	// test DailyTrigger
	dailyTriggerDef := taskService.NewTaskDefinition()
	dailyTriggerDef.AddExecAction("calc.exe", "", "", "")
	dailyTriggerDef.AddDailyTrigger(EveryDay, defaultPeriod, time.Now())
	_, _, err = taskService.CreateTask("\\Taskmaster\\DailyTrigger", dailyTriggerDef, true)
	if err != nil {
		t.Error(err)
	}

	// test EventTrigger
	eventTriggerDef := taskService.NewTaskDefinition()
	eventTriggerDef.AddExecAction("calc.exe", "", "", "")
	subscription := "<QueryList> <Query Id='1'> <Select Path='System'>*[System/Level=2]</Select></Query></QueryList>"
	eventTriggerDef.AddEventTrigger(defaultPeriod, subscription, nil)
	_, _, err = taskService.CreateTask("\\Taskmaster\\EventTrigger", eventTriggerDef, true)
	if err != nil {
		t.Error(err)
	}

	// test IdleTrigger
	idleTriggerDef := taskService.NewTaskDefinition()
	idleTriggerDef.AddExecAction("calc.exe", "", "", "")
	idleTriggerDef.AddIdleTrigger()
	_, _, err = taskService.CreateTask("\\Taskmaster\\IdleTrigger", idleTriggerDef, true)
	if err != nil {
		t.Error(err)
	}

	// test LogonTrigger
	logonTriggerDef := taskService.NewTaskDefinition()
	logonTriggerDef.AddExecAction("calc.exe", "", "", "")
	logonTriggerDef.AddLogonTrigger(defaultPeriod, "")
	_, _, err = taskService.CreateTask("\\Taskmaster\\LogonTrigger", logonTriggerDef, true)
	if err != nil {
		t.Error(err)
	}

	// test MonthlyDOWTrigger
	monthlyDOWTriggerDef := taskService.NewTaskDefinition()
	monthlyDOWTriggerDef.AddExecAction("calc.exe", "", "", "")
	monthlyDOWTriggerDef.AddMonthlyDOWTrigger(Monday|Friday, First, January|February, false, defaultPeriod, time.Now())
	_, _, err = taskService.CreateTask("\\Taskmaster\\MonthlyDOWTrigger", monthlyDOWTriggerDef, true)
	if err != nil {
		t.Error(err)
	}

	// test MonthlyTrigger
	monthlyTriggerDef := taskService.NewTaskDefinition()
	monthlyTriggerDef.AddExecAction("calc.exe", "", "", "")
	monthlyTriggerDef.AddMonthlyTrigger(3, February|March, defaultPeriod, time.Now())
	_, _, err = taskService.CreateTask("\\Taskmaster\\MonthlyTrigger", monthlyTriggerDef, true)
	if err != nil {
		t.Error(err)
	}

	// test RegistrationTrigger
	registrationTriggerDef := taskService.NewTaskDefinition()
	registrationTriggerDef.AddExecAction("calc.exe", "", "", "")
	registrationTriggerDef.AddRegistrationTrigger(defaultPeriod)
	_, _, err = taskService.CreateTask("\\Taskmaster\\RegistrationTrigger", registrationTriggerDef, true)
	if err != nil {
		t.Error(err)
	}

	// test SessionStateChangeTrigger
	sessionStateChangeTriggerDef := taskService.NewTaskDefinition()
	sessionStateChangeTriggerDef.AddExecAction("calc.exe", "", "", "")
	sessionStateChangeTriggerDef.AddSessionStateChangeTrigger("", TASK_SESSION_LOCK, defaultPeriod)
	_, _, err = taskService.CreateTask("\\Taskmaster\\SessionStateChangeTrigger", sessionStateChangeTriggerDef, true)
	if err != nil {
		t.Error(err)
	}

	// test TimeTrigger
	timeTriggerDef := taskService.NewTaskDefinition()
	timeTriggerDef.AddExecAction("calc.exe", "", "", "")
	timeTriggerDef.AddTimeTrigger(defaultPeriod, time.Now())
	_, _, err = taskService.CreateTask("\\Taskmaster\\TimeTrigger", timeTriggerDef, true)
	if err != nil {
		t.Error(err)
	}

	// test WeeklyTrigger
	weeklyTriggerDef := taskService.NewTaskDefinition()
	weeklyTriggerDef.AddExecAction("calc.exe", "", "", "")
	weeklyTriggerDef.AddWeeklyTrigger(Tuesday|Thursday, EveryOtherWeek, defaultPeriod, time.Now())
	_, _, err = taskService.CreateTask("\\Taskmaster\\WeeklyTrigger", weeklyTriggerDef, true)
	if err != nil {
		t.Error(err)
	}

	// test trying to create task where a task at the same path already exists and the 'overwrite' is set to false
	_, taskCreated, err := taskService.CreateTask("\\Taskmaster\\TimeTrigger", timeTriggerDef, false)
	if err != nil {
		t.Error(err)
	}
	if taskCreated {
		t.Error("task shouldn't have been created")
	}
}

func TestUpdateTask(t *testing.T) {
	taskService, err := Connect()
	if err != nil {
		t.Error(err)
	}
	testTask := createTestTask(taskService)
	defer taskService.Disconnect()

	testTask.Definition.RegistrationInfo.Author = "Big Chungus"
	_, err = taskService.UpdateTask("\\Taskmaster\\TestTask", testTask.Definition)
	if err != nil {
		t.Error(err)
	}

	testTask, err = taskService.GetRegisteredTask("\\Taskmaster\\TestTask")
	if err != nil {
		t.Error(err)
	}
	if testTask == nil {
		t.Error("TestTask task should exist")
	}
	if testTask.Definition.RegistrationInfo.Author != "Big Chungus" {
		t.Error("task was not updated")
	}
}

func TestGetRegisteredTasks(t *testing.T) {
	taskService, err := Connect()
	if err != nil {
		t.Error(err)
	}
	defer taskService.Disconnect()

	_, err = taskService.GetRegisteredTasks()
	if err != nil {
		t.Error(err)
	}
}

func TestGetTaskFolders(t *testing.T) {
	taskService, err := Connect()
	if err != nil {
		t.Error(err)
	}
	defer taskService.Disconnect()

	_, err = taskService.GetTaskFolders()
	if err != nil {
		t.Error(err)
	}
}

func TestDeleteTask(t *testing.T) {
	taskService, err := Connect()
	if err != nil {
		t.Error(err)
	}
	createTestTask(taskService)
	defer taskService.Disconnect()

	err = taskService.DeleteTask("\\Taskmaster\\TestTask")
	if err != nil {
		t.Error(err)
	}

	deletedTask, err := taskService.GetRegisteredTask("\\Taskmaster\\TestTask")
	if err != nil {
		t.Error(err)
	}
	if deletedTask != nil {
		t.Error("task should be deleted")
	}
}

func TestDeleteFolder(t *testing.T) {
	taskService, err := Connect()
	if err != nil {
		t.Error(err)
	}
	createTestTask(taskService)
	defer taskService.Disconnect()

	var folderDeleted bool
	folderDeleted, err = taskService.DeleteFolder("\\Taskmaster", false)
	if err != nil {
		t.Error(err)
	}
	if folderDeleted == true {
		t.Error("folder shouldn't have been deleted")
	}

	folderDeleted, err = taskService.DeleteFolder("\\Taskmaster", true)
	if err != nil {
		t.Error(err)
	}
	if folderDeleted == false {
		t.Error("folder should have been deleted")
	}

	tasks, err := taskService.GetRegisteredTasks()
	if err != nil {
		t.Error(err)
	}
	taskmasterFolder, err := taskService.GetTaskFolder("\\Taskmaster")
	if err != nil {
		t.Error(err)
	}
	if taskmasterFolder != nil {
		t.Error("folder shouldn't exist")
	}
	for _, task := range tasks {
		if strings.Split(task.Path, "\\")[1] == "Taskmaster" {
			t.Error("task should've been deleted")
		}
	}
}

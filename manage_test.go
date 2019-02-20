// +build windows

package taskmaster

import (
	"strings"
	"testing"
	"time"
)

func TestLocalConnect(t *testing.T) {
	var taskService TaskService
	err := taskService.Connect("", "", "", "")
	if err != nil {
		t.Error(err)
	}
	taskService.Cleanup()
}

func TestCreateTask(t *testing.T) {
	var err error
	var taskService TaskService
	err = taskService.Connect("", "", "", "")
	if err != nil {
		t.Error(err)
	}
	defer taskService.Cleanup()

	// test ExecAction
	execTaskDef := taskService.NewTaskDefinition()
	execTaskDef.AddExecAction("calc.exe", "", "", "")
	//newTaskDef.AddTimeTrigger("", "", time.Now().Add(5*time.Second), time.Time{}, "", "", "", false, true)

	_, err = taskService.CreateTask("\\Taskmaster\\ExecAction", execTaskDef, "", "", execTaskDef.Principal.LogonType, true)
	if err != nil {
		t.Error(err)
	}

	// test ComHandlerAction
	comHandlerDef := taskService.NewTaskDefinition()
	comHandlerDef.AddComHandlerAction("{F0001111-0000-0000-0000-0000FEEDACDC}", "", "")

	_, err = taskService.CreateTask("\\Taskmaster\\ComHandlerAction", comHandlerDef, "", "", comHandlerDef.Principal.LogonType, true)
	if err != nil {
		t.Error(err)
	}

	// test BootTrigger
	bootTriggerDef := taskService.NewTaskDefinition()
	bootTriggerDef.AddExecAction("calc.exe", "", "", "")
	bootTriggerDef.AddBootTrigger("", "", time.Now(), time.Time{}, "", "", "", false, true)
	_, err = taskService.CreateTask("\\Taskmaster\\BootTrigger", bootTriggerDef, "", "", bootTriggerDef.Principal.LogonType, true)
	if err != nil {
		t.Error(err)
	}

	// test DailyTrigger
	dailyTriggerDef := taskService.NewTaskDefinition()
	dailyTriggerDef.AddExecAction("calc.exe", "", "", "")
	dailyTriggerDef.AddDailyTrigger(EveryDay, "", "", time.Now(), time.Time{}, "", "", "", false, true)
	_, err = taskService.CreateTask("\\Taskmaster\\DailyTrigger", dailyTriggerDef, "", "", dailyTriggerDef.Principal.LogonType, true)
	if err != nil {
		t.Error(err)
	}

	// test EventTrigger
	eventTriggerDef := taskService.NewTaskDefinition()
	eventTriggerDef.AddExecAction("calc.exe", "", "", "")
	subscription := "<QueryList> <Query Id='1'> <Select Path='System'>*[System/Level=2]</Select></Query></QueryList>"
	eventTriggerDef.AddEventTrigger("", subscription, nil, "", time.Now(), time.Time{}, "", "", "", false, true)
	_, err = taskService.CreateTask("\\Taskmaster\\EventTrigger", eventTriggerDef, "", "", eventTriggerDef.Principal.LogonType, true)
	if err != nil {
		t.Error(err)
	}

	// test IdleTrigger
	idleTriggerDef := taskService.NewTaskDefinition()
	idleTriggerDef.AddExecAction("calc.exe", "", "", "")
	idleTriggerDef.AddIdleTrigger("", time.Now(), time.Time{}, "", "", "", false, true)
	_, err = taskService.CreateTask("\\Taskmaster\\IdleTrigger", idleTriggerDef, "", "", idleTriggerDef.Principal.LogonType, true)
	if err != nil {
		t.Error(err)
	}

	// test LogonTrigger
	logonTriggerDef := taskService.NewTaskDefinition()
	logonTriggerDef.AddExecAction("calc.exe", "", "", "")
	logonTriggerDef.AddLogonTrigger("", "", "", time.Now(), time.Time{}, "", "", "", false, true)
	_, err = taskService.CreateTask("\\Taskmaster\\LogonTrigger", logonTriggerDef, "", "", logonTriggerDef.Principal.LogonType, true)
	if err != nil {
		t.Error(err)
	}

	// test MonthlyDOWTrigger
	monthlyDOWTriggerDef := taskService.NewTaskDefinition()
	monthlyDOWTriggerDef.AddExecAction("calc.exe", "", "", "")
	monthlyDOWTriggerDef.AddMonthlyDOWTrigger(Monday|Friday, First, January|February, false, "", "", time.Now(), time.Time{}, "", "", "", false, true)
	_, err = taskService.CreateTask("\\Taskmaster\\MonthlyDOWTrigger", monthlyDOWTriggerDef, "", "", monthlyDOWTriggerDef.Principal.LogonType, true)
	if err != nil {
		t.Error(err)
	}

	// test MonthlyTrigger
	monthlyTriggerDef := taskService.NewTaskDefinition()
	monthlyTriggerDef.AddExecAction("calc.exe", "", "", "")
	monthlyTriggerDef.AddMonthlyTrigger(3, February|March, "", "", time.Now(), time.Time{}, "", "", "", false, true)
	_, err = taskService.CreateTask("\\Taskmaster\\MonthlyTrigger", monthlyTriggerDef, "", "", monthlyTriggerDef.Principal.LogonType, true)
	if err != nil {
		t.Error(err)
	}

	// test RegistrationTrigger
	registrationTriggerDef := taskService.NewTaskDefinition()
	registrationTriggerDef.AddExecAction("calc.exe", "", "", "")
	registrationTriggerDef.AddRegistrationTrigger("", "", time.Now(), time.Time{}, "", "", "", false, true)
	_, err = taskService.CreateTask("\\Taskmaster\\RegistrationTrigger", registrationTriggerDef, "", "", registrationTriggerDef.Principal.LogonType, true)
	if err != nil {
		t.Error(err)
	}

	// test SessionStateChangeTrigger
	sessionStateChangeTriggerDef := taskService.NewTaskDefinition()
	sessionStateChangeTriggerDef.AddExecAction("calc.exe", "", "", "")
	sessionStateChangeTriggerDef.AddSessionStateChangeTrigger("", TASK_SESSION_LOCK, "", "", time.Now(), time.Time{}, "", "", "", false, true)
	_, err = taskService.CreateTask("\\Taskmaster\\SessionStateChangeTrigger", sessionStateChangeTriggerDef, "", "", sessionStateChangeTriggerDef.Principal.LogonType, true)
	if err != nil {
		t.Error(err)
	}

	// test TimeTrigger
	timeTriggerDef := taskService.NewTaskDefinition()
	timeTriggerDef.AddExecAction("calc.exe", "", "", "")
	timeTriggerDef.AddTimeTrigger("", "", time.Now(), time.Time{}, "", "", "", false, true)
	_, err = taskService.CreateTask("\\Taskmaster\\TimeTrigger", timeTriggerDef, "", "", timeTriggerDef.Principal.LogonType, true)
	if err != nil {
		t.Error(err)
	}

	// test WeeklyTrigger
	weeklyTriggerDef := taskService.NewTaskDefinition()
	weeklyTriggerDef.AddExecAction("calc.exe", "", "", "")
	weeklyTriggerDef.AddWeeklyTrigger(Tuesday|Thursday, EveryOtherWeek, "", "", time.Now(), time.Time{}, "", "", "", false, true)
	_, err = taskService.CreateTask("\\Taskmaster\\WeeklyTrigger", weeklyTriggerDef, "", "", weeklyTriggerDef.Principal.LogonType, true)
	if err != nil {
		t.Error(err)
	}

	// test trying to create task where a task at the same path already exists and the 'overwrite' is set to false
	taskCreated, err := taskService.CreateTask("\\Taskmaster\\WeeklyTrigger", timeTriggerDef, "", "", timeTriggerDef.Principal.LogonType, false)
	if err != nil {
		t.Error(err)
	}
	if taskCreated {
		t.Error("task shouldn't have been created")
	}
}

func TestUpdateTask(t *testing.T) {
	var err error
	var taskService TaskService
	err = taskService.Connect("", "", "", "")
	if err != nil {
		t.Error(err)
	}
	defer taskService.Cleanup()

	var task *RegisteredTask
	task, err = taskService.GetRegisteredTask("\\Taskmaster\\WeeklyTrigger")
	if err != nil {
		t.Error(err)
	}
	if task == nil {
		t.Error("WeeklyTrigger task should exist")
	}

	task.Definition.RegistrationInfo.Author = "Big Chungus"
	err = taskService.UpdateTask("\\Taskmaster\\WeeklyTrigger", task.Definition, "", "", task.Definition.Principal.LogonType)
	if err != nil {
		t.Error(err)
	}

	task, err = taskService.GetRegisteredTask("\\Taskmaster\\WeeklyTrigger")
	if err != nil {
		t.Error(err)
	}
	if task == nil {
		t.Error("WeeklyTrigger task should exist")
	}
	if task.Definition.RegistrationInfo.Author != "Big Chungus" {
		t.Error("task was not updated")
	}
}

func TestDeleteTask(t *testing.T) {
	var err error
	var taskService TaskService
	err = taskService.Connect("", "", "", "")
	if err != nil {
		t.Error(err)
	}
	defer taskService.Cleanup()

	err = taskService.DeleteTask("\\Taskmaster\\WeeklyTrigger")
	if err != nil {
		t.Error(err)
	}

	deletedTask, err := taskService.GetRegisteredTask("\\Taskmaster\\WeeklyTrigger")
	if err != nil {
		t.Error(err)
	}
	if deletedTask != nil {
		t.Error("task should be deleted")
	}
}

func TestDeleteFolder(t *testing.T) {
	var err error
	var taskService TaskService
	err = taskService.Connect("", "", "", "")
	if err != nil {
		t.Error(err)
	}
	defer taskService.Cleanup()

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

	err = taskService.GetRegisteredTasks()
	if err != nil {
		t.Error(err)
	}
	for _, taskFolder := range taskService.RootFolder.SubFolders {
		if taskFolder.Name == "Taskmaster" {
			t.Error("folder shouldn't exist")
		}
	}
	for _, task := range taskService.RegisteredTasks {
		if strings.Split(task.Path, "\\")[1] == "Taskmaster" {
			t.Error("task should've been deleted")
		}
	}
}

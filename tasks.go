package taskmaster

import (
	"errors"
	"strings"
	"time"

	ole "github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

const (
	TASK_ENUM_HIDDEN = 1
)

var taskDateFormat = "2006-01-02T15:04:05.0000000"

func (t *TaskService) initialize() error {
	var err error

	err = ole.CoInitialize(0)
	if err != nil {
		return err
	}

	schedClassID, err := ole.ClassIDFrom("Schedule.Service.1")
	if err != nil {
		return err
	}
	taskSchedulerObj, err := ole.CreateInstance(schedClassID, nil)
	if err != nil {
		return err
	}
	if taskSchedulerObj == nil {
		return errors.New("Could not create ITaskService object")
	}
	defer taskSchedulerObj.Release()

	tskSchdlr := taskSchedulerObj.MustQueryInterface(ole.IID_IDispatch)
	t.taskServiceObj = tskSchdlr
	t.isInitialized = true

	return nil
}

// Connect connects to the local Task Scheduler service. This function
// has to be run before any other functions in taskmaster can be used
func (t *TaskService) Connect() error {
	if !t.isInitialized {
		err := t.initialize()
		if err != nil {
			return err
		}
	}

	_, err := oleutil.CallMethod(t.taskServiceObj, "Connect")
	if err != nil {
		return err
	}

	rootFolderObj := oleutil.MustCallMethod(t.taskServiceObj, "GetFolder", "\\").ToIDispatch()
	rootFolder := RootFolder{
		folderObj: rootFolderObj,
		TaskFolder: TaskFolder{
			Name: "\\",
			Path: "\\",
		},
	}
	t.RootFolder = rootFolder

	t.isConnected = true

	return nil
}

// Cleanup frees all the Task Scheduler COM objects that have been created.
// If this function is not called before the parent program terminates,
// memory leaks will occur
func (t *TaskService) Cleanup() {
	for _, runningTask := range t.RunningTasks {
		runningTask.taskObj.Release()
	}

	if t.RootFolder.folderObj != nil {
		t.RootFolder.folderObj.Release()
	}

	for _, registeredTask := range t.RegisteredTasks {
		registeredTask.taskObj.Release()
	}

	t.taskServiceObj.Release()
	ole.CoUninitialize()
	t.isInitialized = false
	t.isConnected = false
}

// GetRunningTasks enumerates the Task Scheduler database for all currently running tasks
func (t *TaskService) GetRunningTasks() error {
	var err error
	t.RunningTasks = nil

	runningTasks := oleutil.MustCallMethod(t.taskServiceObj, "GetRunningTasks", TASK_ENUM_HIDDEN).ToIDispatch()
	defer runningTasks.Release()
	err = oleutil.ForEach(runningTasks, func(v *ole.VARIANT) error {
		task := v.ToIDispatch()

		runningTask := parseRunningTask(task)
		t.RunningTasks = append(t.RunningTasks, &runningTask)

		return nil
	})
	if err != nil {
		return err
	}

	return nil
}

// GetRegisteredTasks enumerates the Task Scheduler database for all currently registered tasks
func (t *TaskService) GetRegisteredTasks() error {
	var err error
	t.RegisteredTasks = nil

	// get tasks from root folder
	rootTaskCollection := oleutil.MustCallMethod(t.RootFolder.folderObj, "GetTasks", TASK_ENUM_HIDDEN).ToIDispatch()
	defer rootTaskCollection.Release()
	err = oleutil.ForEach(rootTaskCollection, func(v *ole.VARIANT) error {
		task := v.ToIDispatch()

		registeredTask, err := parseRegisteredTask(task)
		if err != nil {
			return err
		}
		t.RegisteredTasks = append(t.RegisteredTasks, &registeredTask)
		t.RootFolder.RegisteredTasks = append(t.RootFolder.RegisteredTasks, &registeredTask)

		return nil
	})
	if err != nil {
		return err
	}

	taskFolderList := oleutil.MustCallMethod(t.RootFolder.folderObj, "GetFolders", 0).ToIDispatch()
	defer taskFolderList.Release()

	// recursively enumerate folders and tasks
	var initEnumTaskFolders func(*TaskFolder) func(*ole.VARIANT) error
	initEnumTaskFolders = func(parentFolder *TaskFolder) func(*ole.VARIANT) error {
		var enumTaskFolders func(*ole.VARIANT) error
		enumTaskFolders = func(v *ole.VARIANT) error {
			taskFolder := v.ToIDispatch()
			defer taskFolder.Release()

			name := oleutil.MustGetProperty(taskFolder, "Name").ToString()
			path := oleutil.MustGetProperty(taskFolder, "Path").ToString()
			taskCollection := oleutil.MustCallMethod(taskFolder, "GetTasks", TASK_ENUM_HIDDEN).ToIDispatch()
			defer taskCollection.Release()

			taskSubFolder := &TaskFolder{
				Name: name,
				Path: path,
			}

			var err error
			err = oleutil.ForEach(taskCollection, func(v *ole.VARIANT) error {
				task := v.ToIDispatch()

				registeredTask, err := parseRegisteredTask(task)
				if err != nil {
					return err
				}
				t.RegisteredTasks = append(t.RegisteredTasks, &registeredTask)
				taskSubFolder.RegisteredTasks = append(taskSubFolder.RegisteredTasks, &registeredTask)

				return nil
			})
			if err != nil {
				return err
			}

			parentFolder.SubFolders = append(parentFolder.SubFolders, taskSubFolder)

			taskFolderList := oleutil.MustCallMethod(taskFolder, "GetFolders", 0).ToIDispatch()
			defer taskFolderList.Release()

			err = oleutil.ForEach(taskFolderList, initEnumTaskFolders(taskSubFolder))
			if err != nil {
				return err
			}

			return nil
		}

		return enumTaskFolders
	}

	err = oleutil.ForEach(taskFolderList, initEnumTaskFolders(&t.RootFolder.TaskFolder))
	if err != nil {
		return err
	}

	return nil
}

func parseRunningTask(task *ole.IDispatch) RunningTask {
	currentAction := oleutil.MustGetProperty(task, "CurrentAction").ToString()
	enginePID := int(oleutil.MustGetProperty(task, "EnginePid").Val)
	instanceGUID := oleutil.MustGetProperty(task, "InstanceGuid").ToString()
	name := oleutil.MustGetProperty(task, "Name").ToString()
	path := oleutil.MustGetProperty(task, "Path").ToString()
	state := int(oleutil.MustGetProperty(task, "State").Val)

	runningTask := RunningTask{
		taskObj:       task,
		CurrentAction: currentAction,
		EnginePID:     enginePID,
		InstanceGUID:  instanceGUID,
		Name:          name,
		Path:          path,
		State:         state,
	}

	return runningTask
}

func parseRegisteredTask(task *ole.IDispatch) (RegisteredTask, error) {
	var err error

	name := oleutil.MustGetProperty(task, "Name").ToString()
	path := oleutil.MustGetProperty(task, "Path").ToString()
	enabled := oleutil.MustGetProperty(task, "Enabled").Value().(bool)
	state := int(oleutil.MustGetProperty(task, "State").Val)
	missedRuns := int(oleutil.MustGetProperty(task, "NumberOfMissedRuns").Val)
	nextRunTime := oleutil.MustGetProperty(task, "NextRunTime").Value().(time.Time)
	lastRunTime := oleutil.MustGetProperty(task, "LastRunTime").Value().(time.Time)
	lastTaskResult := int(oleutil.MustGetProperty(task, "LastTaskResult").Val)

	definition := oleutil.MustGetProperty(task, "Definition").ToIDispatch()
	defer definition.Release()
	actions := oleutil.MustGetProperty(definition, "Actions").ToIDispatch()
	defer actions.Release()
	context := oleutil.MustGetProperty(actions, "Context").ToString()

	var taskActions []Action
	err = oleutil.ForEach(actions, func(v *ole.VARIANT) error {
		action := v.ToIDispatch()
		defer action.Release()

		taskAction, err := parseTaskAction(action)
		if err != nil {
			return err
		}

		taskActions = append(taskActions, taskAction)

		return nil
	})
	if err != nil {
		return RegisteredTask{}, err
	}

	principal := oleutil.MustGetProperty(definition, "Principal").ToIDispatch()
	defer principal.Release()
	taskPrincipal := parsePrincipal(principal)

	regInfo := oleutil.MustGetProperty(definition, "RegistrationInfo").ToIDispatch()
	defer regInfo.Release()
	registrationInfo := parseRegistrationInfo(regInfo)

	settings := oleutil.MustGetProperty(definition, "Settings").ToIDispatch()
	defer settings.Release()
	taskSettings := parseTaskSettings(settings)

	triggers := oleutil.MustGetProperty(definition, "Triggers").ToIDispatch()
	defer triggers.Release()

	var taskTriggers []Trigger
	err = oleutil.ForEach(triggers, func(v *ole.VARIANT) error {
		trigger := v.ToIDispatch()
		defer trigger.Release()

		taskTrigger, err := parseTaskTrigger(trigger)
		if err != nil {
			return err
		}
		taskTriggers = append(taskTriggers, taskTrigger)

		return nil
	})
	if err != nil {
		return RegisteredTask{}, err
	}

	taskDef := Definition{
		Actions:          taskActions,
		Context:          context,
		Principal:        taskPrincipal,
		Settings:         taskSettings,
		RegistrationInfo: registrationInfo,
		Triggers:         taskTriggers,
	}

	RegisteredTask := RegisteredTask{
		taskObj:        task,
		Name:           name,
		Path:           path,
		Definition:     taskDef,
		Enabled:        enabled,
		State:          state,
		MissedRuns:     missedRuns,
		NextRunTime:    nextRunTime,
		LastRunTime:    lastRunTime,
		LastTaskResult: lastTaskResult,
	}

	return RegisteredTask, nil
}

func parseTaskAction(action *ole.IDispatch) (Action, error) {
	id := oleutil.MustGetProperty(action, "Id").ToString()
	actionType := int(oleutil.MustGetProperty(action, "Type").Val)

	switch actionType {
	case TASK_ACTION_EXEC:
		args := oleutil.MustGetProperty(action, "Arguments").ToString()
		path := oleutil.MustGetProperty(action, "Path").ToString()
		workingDir := oleutil.MustGetProperty(action, "WorkingDirectory").ToString()

		execAction := ExecAction{
			TaskAction: TaskAction{
				ID: id,
				TypeHolder: TypeHolder{
					Type: actionType,
				},
			},
			Path:       path,
			Args:       args,
			WorkingDir: workingDir,
		}

		return execAction, nil
	case TASK_ACTION_COM_HANDLER:
		classID := oleutil.MustGetProperty(action, "ClassId").ToString()
		data := oleutil.MustGetProperty(action, "Data").ToString()

		comHandlerAction := ComHandlerAction{
			TaskAction: TaskAction{
				ID: id,
				TypeHolder: TypeHolder{
					Type: actionType,
				},
			},
			ClassID: classID,
			Data:    data,
		}

		return comHandlerAction, nil
	default:
		return nil, errors.New("unsupported IAction type")
	}
}

func parsePrincipal(principleObj *ole.IDispatch) Principal {
	name := oleutil.MustGetProperty(principleObj, "DisplayName").ToString()
	groupID := oleutil.MustGetProperty(principleObj, "GroupId").ToString()
	id := oleutil.MustGetProperty(principleObj, "Id").ToString()
	logonType := int(oleutil.MustGetProperty(principleObj, "LogonType").Val)
	runLevel := int(oleutil.MustGetProperty(principleObj, "RunLevel").Val)
	userID := oleutil.MustGetProperty(principleObj, "UserId").ToString()

	principle := Principal{
		Name:      name,
		GroupID:   groupID,
		ID:        id,
		LogonType: logonType,
		RunLevel:  runLevel,
		UserID:    userID,
	}

	return principle
}

func parseRegistrationInfo(regInfo *ole.IDispatch) RegistrationInfo {
	author := oleutil.MustGetProperty(regInfo, "Author").ToString()
	date := oleutil.MustGetProperty(regInfo, "Date").ToString()
	description := oleutil.MustGetProperty(regInfo, "Description").ToString()
	documentation := oleutil.MustGetProperty(regInfo, "Documentation").ToString()
	securityDescriptor := oleutil.MustGetProperty(regInfo, "SecurityDescriptor").ToString()
	source := oleutil.MustGetProperty(regInfo, "Source").ToString()
	uri := oleutil.MustGetProperty(regInfo, "URI").ToString()
	version := oleutil.MustGetProperty(regInfo, "Version").ToString()

	registrationInfo := RegistrationInfo{
		Author:             author,
		Date:               date,
		Description:        description,
		Documentation:      documentation,
		SecurityDescriptor: securityDescriptor,
		Source:             source,
		URI:                uri,
		Version:            version,
	}

	return registrationInfo
}

func parseTaskSettings(settings *ole.IDispatch) TaskSettings {
	allowDemandStart := oleutil.MustGetProperty(settings, "AllowDemandStart").Value().(bool)
	allowHardTerminate := oleutil.MustGetProperty(settings, "AllowHardTerminate").Value().(bool)
	compatibility := int(oleutil.MustGetProperty(settings, "Compatibility").Val)
	deleteExpiredTaskAfter := oleutil.MustGetProperty(settings, "DeleteExpiredTaskAfter").ToString()
	dontStartOnBatteries := oleutil.MustGetProperty(settings, "DisallowStartIfOnBatteries").Value().(bool)
	enabled := oleutil.MustGetProperty(settings, "Enabled").Value().(bool)
	timeLimit := oleutil.MustGetProperty(settings, "ExecutionTimeLimit").ToString()
	hidden := oleutil.MustGetProperty(settings, "Hidden").Value().(bool)

	idleSettings := oleutil.MustGetProperty(settings, "IdleSettings").ToIDispatch()
	defer idleSettings.Release()
	idleDuration := oleutil.MustGetProperty(idleSettings, "IdleDuration").ToString()
	restartOnIdle := oleutil.MustGetProperty(idleSettings, "RestartOnIdle").Value().(bool)
	stopOnIdleEnd := oleutil.MustGetProperty(idleSettings, "StopOnIdleEnd").Value().(bool)
	waitTimeOut := oleutil.MustGetProperty(idleSettings, "WaitTimeout").ToString()

	multipleInstances := int(oleutil.MustGetProperty(settings, "MultipleInstances").Val)

	networkSettings := oleutil.MustGetProperty(settings, "NetworkSettings").ToIDispatch()
	defer networkSettings.Release()
	id := oleutil.MustGetProperty(networkSettings, "Id").ToString()
	name := oleutil.MustGetProperty(networkSettings, "Name").ToString()

	priority := int(oleutil.MustGetProperty(settings, "Priority").Val)
	restartCount := int(oleutil.MustGetProperty(settings, "RestartCount").Val)
	restartInterval := oleutil.MustGetProperty(settings, "RestartInterval").ToString()
	runOnlyIfIdle := oleutil.MustGetProperty(settings, "RunOnlyIfIdle").Value().(bool)
	runOnlyIfNetworkAvalible := oleutil.MustGetProperty(settings, "RunOnlyIfNetworkAvailable").Value().(bool)
	startWhenAvalible := oleutil.MustGetProperty(settings, "StartWhenAvailable").Value().(bool)
	stopIfGoingOnBatteries := oleutil.MustGetProperty(settings, "StopIfGoingOnBatteries").Value().(bool)
	wakeToRun := oleutil.MustGetProperty(settings, "WakeToRun").Value().(bool)

	idleTaskSettings := IdleSettings{
		IdleDuration:  idleDuration,
		RestartOnIdle: restartOnIdle,
		StopOnIdleEnd: stopOnIdleEnd,
		WaitTimeout:   waitTimeOut,
	}

	networkTaskSettings := NetworkSettings{
		ID:   id,
		Name: name,
	}

	taskSettings := TaskSettings{
		AllowDemandStart:         allowDemandStart,
		AllowHardTerminate:       allowHardTerminate,
		Compatibility:            compatibility,
		DeleteExpiredTaskAfter:   deleteExpiredTaskAfter,
		DontStartOnBatteries:     dontStartOnBatteries,
		Enabled:                  enabled,
		TimeLimit:                timeLimit,
		Hidden:                   hidden,
		IdleSettings:             idleTaskSettings,
		MultipleInstances:        multipleInstances,
		NetworkSettings:          networkTaskSettings,
		Priority:                 priority,
		RestartCount:             restartCount,
		RestartInterval:          restartInterval,
		RunOnlyIfIdle:            runOnlyIfIdle,
		RunOnlyIfNetworkAvalible: runOnlyIfNetworkAvalible,
		StartWhenAvalible:        startWhenAvalible,
		StopIfGoingOnBatteries:   stopIfGoingOnBatteries,
		WakeToRun:                wakeToRun,
	}

	return taskSettings
}

func parseTaskTrigger(trigger *ole.IDispatch) (Trigger, error) {
	enabled := oleutil.MustGetProperty(trigger, "Enabled").Value().(bool)
	endBoundary := oleutil.MustGetProperty(trigger, "EndBoundary").ToString()
	executionTimeLimit := oleutil.MustGetProperty(trigger, "ExecutionTimeLimit").ToString()
	id := oleutil.MustGetProperty(trigger, "Id").ToString()

	repetition := oleutil.MustGetProperty(trigger, "Repetition").ToIDispatch()
	defer repetition.Release()
	duration := oleutil.MustGetProperty(repetition, "Duration").ToString()
	interval := oleutil.MustGetProperty(repetition, "Interval").ToString()
	stopAtDurationEnd := oleutil.MustGetProperty(repetition, "StopAtDurationEnd").Value().(bool)

	startBoundary := oleutil.MustGetProperty(trigger, "StartBoundary").ToString()
	triggerType := int(oleutil.MustGetProperty(trigger, "Type").Val)

	taskTriggerObj := TaskTrigger{
		Enabled:            enabled,
		EndBoundary:        endBoundary,
		ExecutionTimeLimit: executionTimeLimit,
		ID:                 id,
		RepetitionPattern: RepetitionPattern{
			Duration:          duration,
			Interval:          interval,
			StopAtDurationEnd: stopAtDurationEnd,
		},
		StartBoundary: startBoundary,
		TypeHolder: TypeHolder{
			Type: triggerType,
		},
	}

	switch triggerType {
	case TASK_TRIGGER_BOOT:
		delay := oleutil.MustGetProperty(trigger, "Delay").ToString()

		bootTrigger := BootTrigger{
			TaskTrigger: taskTriggerObj,
			Delay:       delay,
		}

		return bootTrigger, nil
	case TASK_TRIGGER_DAILY:
		daysInterval := int(oleutil.MustGetProperty(trigger, "DaysInterval").Val)
		randomDelay := oleutil.MustGetProperty(trigger, "RandomDelay").ToString()

		dailyTrigger := DailyTrigger{
			TaskTrigger:  taskTriggerObj,
			DaysInterval: daysInterval,
			RandomDelay:  randomDelay,
		}

		return dailyTrigger, nil
	case TASK_TRIGGER_EVENT:
		delay := oleutil.MustGetProperty(trigger, "Delay").ToString()
		subscription := oleutil.MustGetProperty(trigger, "Subscription").ToString()
		valueQueriesObj := oleutil.MustGetProperty(trigger, "ValueQueries").ToIDispatch()
		defer valueQueriesObj.Release()

		valQueryMap := make(map[string]string)
		err := oleutil.ForEach(valueQueriesObj, func(v *ole.VARIANT) error {
			valueQuery := v.ToIDispatch()
			defer valueQuery.Release()

			name := oleutil.MustGetProperty(valueQuery, "Name").ToString()
			value := oleutil.MustGetProperty(valueQuery, "Value").ToString()

			valQueryMap[name] = value

			return nil
		})
		if err != nil {
			return nil, err
		}

		eventTrigger := EventTrigger{
			TaskTrigger:  taskTriggerObj,
			Delay:        delay,
			Subscription: subscription,
			ValueQueries: valQueryMap,
		}

		return eventTrigger, nil
	case TASK_TRIGGER_IDLE:
		idleTrigger := IdleTrigger{
			TaskTrigger: taskTriggerObj,
		}

		return idleTrigger, nil
	case TASK_TRIGGER_LOGON:
		delay := oleutil.MustGetProperty(trigger, "Delay").ToString()
		userID := oleutil.MustGetProperty(trigger, "UserId").ToString()

		logonTrigger := LogonTrigger{
			TaskTrigger: taskTriggerObj,
			Delay:       delay,
			UserID:      userID,
		}

		return logonTrigger, nil
	case TASK_TRIGGER_MONTHLYDOW:
		daysOfWeek := int(oleutil.MustGetProperty(trigger, "DaysOfWeek").Val)
		monthsOfYear := int(oleutil.MustGetProperty(trigger, "MonthsOfYear").Val)
		randomDelay := oleutil.MustGetProperty(trigger, "RandomDelay").ToString()
		runOnLastWeekOfMonth := oleutil.MustGetProperty(trigger, "RunOnLastWeekOfMonth").Value().(bool)
		weeksOfMonth := int(oleutil.MustGetProperty(trigger, "WeeksOfMonth").Val)

		monthlyDOWTrigger := MonthlyDOWTrigger{
			TaskTrigger:          taskTriggerObj,
			DaysOfWeek:           daysOfWeek,
			MonthsOfYear:         monthsOfYear,
			RandomDelay:          randomDelay,
			RunOnLastWeekOfMonth: runOnLastWeekOfMonth,
			WeeksOfMonth:         weeksOfMonth,
		}

		return monthlyDOWTrigger, nil
	case TASK_TRIGGER_MONTHLY:
		daysOfMonth := int(oleutil.MustGetProperty(trigger, "DaysOfMonth").Val)
		monthsOfYear := int(oleutil.MustGetProperty(trigger, "MonthsOfYear").Val)
		randomDelay := oleutil.MustGetProperty(trigger, "RandomDelay").ToString()
		runOnLastWeekOfMonth := oleutil.MustGetProperty(trigger, "RunOnLastDayOfMonth").Value().(bool)

		monthlyTrigger := MonthlyTrigger{
			TaskTrigger:          taskTriggerObj,
			DaysOfMonth:          daysOfMonth,
			MonthsOfYear:         monthsOfYear,
			RandomDelay:          randomDelay,
			RunOnLastWeekOfMonth: runOnLastWeekOfMonth,
		}

		return monthlyTrigger, nil
	case TASK_TRIGGER_REGISTRATION:
		delay := oleutil.MustGetProperty(trigger, "Delay").ToString()

		registrationTrigger := RegistrationTrigger{
			TaskTrigger: taskTriggerObj,
			Delay:       delay,
		}

		return registrationTrigger, nil
	case TASK_TRIGGER_TIME:
		randomDelay := oleutil.MustGetProperty(trigger, "RandomDelay").ToString()

		timetrigger := TimeTrigger{
			TaskTrigger: taskTriggerObj,
			RandomDelay: randomDelay,
		}

		return timetrigger, nil
	case TASK_TRIGGER_WEEKLY:
		daysOfWeek := int(oleutil.MustGetProperty(trigger, "DaysOfWeek").Val)
		randomDelay := oleutil.MustGetProperty(trigger, "RandomDelay").ToString()
		weeksInterval := int(oleutil.MustGetProperty(trigger, "WeeksInterval").Val)

		weeklyTrigger := WeeklyTrigger{
			TaskTrigger:   taskTriggerObj,
			DaysOfWeek:    daysOfWeek,
			RandomDelay:   randomDelay,
			WeeksInterval: weeksInterval,
		}

		return weeklyTrigger, nil
	case TASK_TRIGGER_SESSION_STATE_CHANGE:
		delay := oleutil.MustGetProperty(trigger, "Delay").ToString()
		stateChange := int(oleutil.MustGetProperty(trigger, "StateChange").Val)
		userID := oleutil.MustGetProperty(trigger, "UserId").ToString()

		sessionStateChangeTrigger := SessionStateChangeTrigger{
			TaskTrigger: taskTriggerObj,
			Delay:       delay,
			StateChange: stateChange,
			UserId:      userID,
		}

		return sessionStateChangeTrigger, nil
	case TASK_TRIGGER_CUSTOM_TRIGGER_01:
		customTrigger := CustomTrigger{
			TaskTrigger: taskTriggerObj,
		}

		return customTrigger, nil

	default:
		return nil, errors.New("unsupported ITrigger type")
	}
}

func NewTaskDefinition() Definition {
	newDef := Definition{}

	newDef.Principal.LogonType = TASK_LOGON_INTERACTIVE_TOKEN
	newDef.Principal.RunLevel = TASK_RUNLEVEL_LUA

	newDef.RegistrationInfo.Date = TimeToTaskDate(time.Now())

	newDef.Settings.AllowDemandStart = true
	newDef.Settings.AllowHardTerminate = true
	newDef.Settings.Compatibility = TASK_COMPATIBILITY_V2
	newDef.Settings.DontStartOnBatteries = true
	newDef.Settings.Enabled = true
	newDef.Settings.Hidden = false
	newDef.Settings.IdleSettings.IdleDuration = "PT10M"
	newDef.Settings.IdleSettings.WaitTimeout = "PT1H"
	newDef.Settings.MultipleInstances = TASK_INSTANCES_IGNORE_NEW
	newDef.Settings.Priority = 7
	newDef.Settings.RestartCount = 0
	newDef.Settings.RestartOnIdle = false
	newDef.Settings.RunOnlyIfIdle = false
	newDef.Settings.RunOnlyIfNetworkAvalible = false
	newDef.Settings.StartWhenAvalible = false
	newDef.Settings.StopIfGoingOnBatteries = true
	newDef.Settings.StopOnIdleEnd = true
	newDef.Settings.TimeLimit = "PT72H"
	newDef.Settings.WakeToRun = false

	return newDef
}

func TimeToTaskDate(t time.Time) string {
	return t.Format(taskDateFormat)
}

func TaskDateToTime(s string) (time.Time, error) {
	t, err := time.Parse(taskDateFormat, s)
	if err != nil {
		return time.Time{}, err
	}

	return t, nil
}

func (t *TaskService) CreateTask(path string, newTaskDef Definition, username, password string, logonType int, overwrite bool) (bool, error) {
	var err error

	if path[0] != '\\' {
		return false, errors.New("path must start with root folder '\\'")
	}

	nameIndex := strings.LastIndex(path, "\\")
	folderPath := path[:nameIndex]

	if !t.doesTaskFolderExist(folderPath) {
		oleutil.MustCallMethod(t.RootFolder.folderObj, "CreateFolder", folderPath, "")
	} else {
		_, err = oleutil.CallMethod(t.RootFolder.folderObj, "GetTask", path)
		if err == nil {
			if !overwrite {
				return false, nil
			}
			oleutil.CallMethod(t.RootFolder.folderObj, "DeleteTask", path, 0)
		}
	}

	_, err = t.modifyTask(path, newTaskDef, username, password, logonType, TASK_CREATE)
	if err != nil {
		return false, err
	}

	// TODO: update taskService with possibly new folders and the new task

	return true, nil
}

func (t *TaskService) UpdateTask(path string, newTaskDef Definition, username, password string, logonType int) error {
	var err error

	if path[0] != '\\' {
		return errors.New("path must start with root folder '\\'")
	}

	_, err = oleutil.CallMethod(t.RootFolder.folderObj, "GetTask", path)
	if err == nil {
		return errors.New("registered task doesn't exist")
	}

	newTaskObj, err := t.modifyTask(path, newTaskDef, username, password, logonType, TASK_UPDATE)
	if err != nil {
		return err
	}

	for _, task := range t.RegisteredTasks {
		if task.Path == path {
			newTask, err := parseRegisteredTask(newTaskObj)
			if err != nil {
				return err
			}
			task = &newTask
		}
	}

	return nil
}

func (t *TaskService) modifyTask(path string, newTaskDef Definition, username, password string, logonType int, flags int) (*ole.IDispatch, error) {
	var err error

	newTaskDefObj := oleutil.MustCallMethod(t.taskServiceObj, "NewTask", 0).ToIDispatch()
	defer newTaskDefObj.Release()

	err = fillDefinitionObj(newTaskDef, newTaskDefObj)
	if err != nil {
		return nil, err
	}

	// make sure registration triggers won't trigger themselves
	for _, trigger := range newTaskDef.Triggers {
		if trigger.GetType() == TASK_TRIGGER_REGISTRATION {
			flags |= TASK_IGNORE_REGISTRATION_TRIGGERS
		}
	}

	newTaskObj, err := oleutil.CallMethod(t.RootFolder.folderObj, "RegisterTaskDefinition", path, newTaskDefObj, flags, username, password, logonType, "")
	if err != nil {
		return nil, err
	}

	return newTaskObj.ToIDispatch(), nil
}

func (t *TaskService) doesTaskFolderExist(path string) bool {
	_, err := oleutil.CallMethod(t.taskServiceObj, "GetFolder", path)
	if err != nil {
		return false
	}

	return true
}

func fillDefinitionObj(definition Definition, definitionObj *ole.IDispatch) error {
	var err error

	actionsObj := oleutil.MustGetProperty(definitionObj, "Actions").ToIDispatch()
	defer actionsObj.Release()
	oleutil.MustPutProperty(actionsObj, "Context", definition.Context)
	err = fillActionsObj(definition.Actions, actionsObj)
	if err != nil {
		return err
	}

	oleutil.MustPutProperty(definitionObj, "Data", definition.Data)

	principalObj := oleutil.MustGetProperty(definitionObj, "Principal").ToIDispatch()
	defer principalObj.Release()
	fillPrincipalObj(definition.Principal, principalObj)

	regInfoObj := oleutil.MustGetProperty(definitionObj, "RegistrationInfo").ToIDispatch()
	defer regInfoObj.Release()
	fillRegistrationInfoObj(definition.RegistrationInfo, regInfoObj)

	settingsObj := oleutil.MustGetProperty(definitionObj, "Settings").ToIDispatch()
	defer settingsObj.Release()
	fillTaskSettingsObj(definition.Settings, settingsObj)

	triggersObj := oleutil.MustGetProperty(definitionObj, "Triggers").ToIDispatch()
	defer triggersObj.Release()
	err = fillTaskTriggersObj(definition.Triggers, triggersObj)
	if err != nil {
		return err
	}

	return nil
}

func fillActionsObj(actions []Action, actionsObj *ole.IDispatch) error {
	for _, action := range actions {
		actionType := action.GetType()
		if !checkActionType(actionType) {
			return errors.New("invalid action type")
		}

		actionObj := oleutil.MustCallMethod(actionsObj, "Create", actionType).ToIDispatch()
		actionObj.Release()
		oleutil.MustPutProperty(actionObj, "Id", action.GetID())

		switch actionType {
		case TASK_ACTION_EXEC:
			execAction := action.(ExecAction)
			exeActionObj := actionObj.MustQueryInterface(ole.NewGUID("{4c3d624d-fd6b-49a3-b9b7-09cb3cd3f047}"))
			defer exeActionObj.Release()

			oleutil.MustPutProperty(exeActionObj, "Arguments", execAction.Args)
			oleutil.MustPutProperty(exeActionObj, "Path", execAction.Path)
			oleutil.MustPutProperty(exeActionObj, "WorkingDirectory", execAction.WorkingDir)
		case TASK_ACTION_COM_HANDLER:
			comHandlerAction := action.(ComHandlerAction)
			comHandlerActionObj := actionObj.MustQueryInterface(ole.NewGUID("{6d2fd252-75c5-4f66-90ba-2a7d8cc3039f}"))
			defer comHandlerActionObj.Release()

			oleutil.MustPutProperty(actionsObj, "ClassId", comHandlerAction.ClassID)
			oleutil.MustPutProperty(actionsObj, "Data", comHandlerAction.Data)
		}
	}

	return nil
}

func checkActionType(actionType int) bool {
	switch actionType {
	case TASK_ACTION_EXEC:
		fallthrough
	case TASK_ACTION_COM_HANDLER:
		fallthrough
	case TASK_ACTION_SEND_EMAIL:
		fallthrough
	case TASK_ACTION_SHOW_MESSAGE:
		return true
	default:
		return false
	}
}

func fillPrincipalObj(principal Principal, principalObj *ole.IDispatch) {
	oleutil.MustPutProperty(principalObj, "DisplayName", principal.Name)
	oleutil.MustPutProperty(principalObj, "GroupId", principal.GroupID)
	oleutil.MustPutProperty(principalObj, "Id", principal.ID)
	oleutil.MustPutProperty(principalObj, "LogonType", principal.LogonType)
	oleutil.MustPutProperty(principalObj, "RunLevel", principal.RunLevel)
	oleutil.MustPutProperty(principalObj, "UserId", principal.UserID)
}

func fillRegistrationInfoObj(regInfo RegistrationInfo, regInfoObj *ole.IDispatch) {
	oleutil.MustPutProperty(regInfoObj, "Author", regInfo.Author)
	oleutil.MustPutProperty(regInfoObj, "Date", regInfo.Date)
	oleutil.MustPutProperty(regInfoObj, "Description", regInfo.Description)
	oleutil.MustPutProperty(regInfoObj, "Documentation", regInfo.Documentation)
	oleutil.MustPutProperty(regInfoObj, "SecurityDescriptor", regInfo.SecurityDescriptor)
	oleutil.MustPutProperty(regInfoObj, "Source", regInfo.Source)
	oleutil.MustPutProperty(regInfoObj, "URI", regInfo.URI)
	oleutil.MustPutProperty(regInfoObj, "Version", regInfo.Version)
}

func fillTaskSettingsObj(settings TaskSettings, settingsObj *ole.IDispatch) {
	oleutil.MustPutProperty(settingsObj, "AllowDemandStart", settings.AllowDemandStart)
	oleutil.MustPutProperty(settingsObj, "AllowHardTerminate", settings.AllowHardTerminate)
	oleutil.MustPutProperty(settingsObj, "Compatibility", settings.Compatibility)
	oleutil.MustPutProperty(settingsObj, "DeleteExpiredTaskAfter", settings.DeleteExpiredTaskAfter)
	oleutil.MustPutProperty(settingsObj, "DisallowStartIfOnBatteries", settings.DontStartOnBatteries)
	oleutil.MustPutProperty(settingsObj, "Enabled", settings.Enabled)
	oleutil.MustPutProperty(settingsObj, "ExecutionTimeLimit", settings.TimeLimit)
	oleutil.MustPutProperty(settingsObj, "Hidden", settings.Hidden)

	idlesettingsObj := oleutil.MustGetProperty(settingsObj, "IdleSettings").ToIDispatch()
	defer idlesettingsObj.Release()
	oleutil.MustPutProperty(idlesettingsObj, "IdleDuration", settings.IdleSettings.IdleDuration)
	oleutil.MustPutProperty(idlesettingsObj, "RestartOnIdle", settings.IdleSettings.RestartOnIdle)
	oleutil.MustPutProperty(idlesettingsObj, "StopOnIdleEnd", settings.IdleSettings.StopOnIdleEnd)
	oleutil.MustPutProperty(idlesettingsObj, "WaitTimeout", settings.IdleSettings.WaitTimeout)

	oleutil.MustPutProperty(settingsObj, "MultipleInstances", settings.MultipleInstances)

	networksettingsObj := oleutil.MustGetProperty(settingsObj, "NetworkSettings").ToIDispatch()
	defer networksettingsObj.Release()
	oleutil.MustPutProperty(networksettingsObj, "Id", settings.NetworkSettings.ID)
	oleutil.MustPutProperty(networksettingsObj, "Name", settings.NetworkSettings.Name)

	oleutil.MustPutProperty(settingsObj, "Priority", settings.Priority)
	oleutil.MustPutProperty(settingsObj, "RestartCount", settings.RestartCount)
	oleutil.MustPutProperty(settingsObj, "RestartInterval", settings.RestartInterval)
	oleutil.MustPutProperty(settingsObj, "RunOnlyIfIdle", settings.RunOnlyIfIdle)
	oleutil.MustPutProperty(settingsObj, "RunOnlyIfNetworkAvailable", settings.RunOnlyIfNetworkAvalible)
	oleutil.MustPutProperty(settingsObj, "StartWhenAvailable", settings.StartWhenAvalible)
	oleutil.MustPutProperty(settingsObj, "StopIfGoingOnBatteries", settings.StopIfGoingOnBatteries)
	oleutil.MustPutProperty(settingsObj, "WakeToRun", settings.WakeToRun)
}

func fillTaskTriggersObj(triggers []Trigger, triggersObj *ole.IDispatch) error {
	for _, trigger := range triggers {
		triggerType := trigger.GetType()
		if !checkTriggerType(triggerType) {
			return errors.New("invalid trigger type")
		}
		triggerObj := oleutil.MustCallMethod(triggersObj, "Create", triggerType).ToIDispatch()
		defer triggerObj.Release()

		oleutil.MustPutProperty(triggerObj, "Enabled", trigger.GetEnabled())
		oleutil.MustPutProperty(triggerObj, "EndBoundary", trigger.GetEndBoundary())
		oleutil.MustPutProperty(triggerObj, "ExecutionTimeLimit", trigger.GetExecutionTimeLimit())
		oleutil.MustPutProperty(triggerObj, "Id", trigger.GetID())

		repetitionObj := oleutil.MustGetProperty(triggerObj, "Repetition").ToIDispatch()
		defer repetitionObj.Release()
		oleutil.MustPutProperty(repetitionObj, "Duration", trigger.GetDuration())
		oleutil.MustPutProperty(repetitionObj, "Interval", trigger.GetInterval())
		oleutil.MustPutProperty(repetitionObj, "StopAtDurationEnd", trigger.GetStopAtDurationEnd())

		oleutil.MustPutProperty(triggerObj, "StartBoundary", trigger.GetStartBoundary())

		switch triggerType {
		case TASK_TRIGGER_BOOT:
			bootTrigger := trigger.(BootTrigger)
			bootTriggerObj := triggerObj.MustQueryInterface(ole.NewGUID("{2a9c35da-d357-41f4-bbc1-207ac1b1f3cb}"))
			defer bootTriggerObj.Release()

			oleutil.MustPutProperty(bootTriggerObj, "Delay", bootTrigger.Delay)
		case TASK_TRIGGER_DAILY:
			dailyTrigger := trigger.(DailyTrigger)
			dailyTriggerObj := triggerObj.MustQueryInterface(ole.NewGUID("{126c5cd8-b288-41d5-8dbf-e491446adc5c}"))
			defer dailyTriggerObj.Release()

			oleutil.MustPutProperty(dailyTriggerObj, "DaysInterval", dailyTrigger.DaysInterval)
			oleutil.MustPutProperty(dailyTriggerObj, "RandomDelay", dailyTrigger.RandomDelay)
		case TASK_TRIGGER_EVENT:
			eventTrigger := trigger.(EventTrigger)
			eventTriggerObj := triggerObj.MustQueryInterface(ole.NewGUID("{d45b0167-9653-4eef-b94f-0732ca7af251}"))
			defer eventTriggerObj.Release()

			oleutil.MustPutProperty(eventTriggerObj, "Delay", eventTrigger.Delay)
			oleutil.MustPutProperty(eventTriggerObj, "Subscription", eventTrigger.Subscription)
			valueQueriesObj := oleutil.MustGetProperty(eventTriggerObj, "ValueQueries").ToIDispatch()
			defer valueQueriesObj.Release()

			for name, value := range eventTrigger.ValueQueries {
				oleutil.MustCallMethod(valueQueriesObj, "Create", name, value)
			}
		case TASK_TRIGGER_IDLE:
			idleTriggerObj := triggerObj.MustQueryInterface(ole.NewGUID("{d537d2b0-9fb3-4d34-9739-1ff5ce7b1ef3}"))
			defer idleTriggerObj.Release()
		case TASK_TRIGGER_LOGON:
			logonTrigger := trigger.(LogonTrigger)
			logonTriggerObj := triggerObj.MustQueryInterface(ole.NewGUID("{72dade38-fae4-4b3e-baf4-5d009af02b1c}"))
			defer logonTriggerObj.Release()

			oleutil.MustPutProperty(logonTriggerObj, "Delay", logonTrigger.Delay)
			oleutil.MustPutProperty(logonTriggerObj, "UserId", logonTrigger.UserID)
		case TASK_TRIGGER_MONTHLYDOW:
			monthlyDOWTrigger := trigger.(MonthlyDOWTrigger)
			monthlyDOWTriggerObj := triggerObj.MustQueryInterface(ole.NewGUID("{77d025a3-90fa-43aa-b52e-cda5499b946a}"))
			defer monthlyDOWTriggerObj.Release()

			oleutil.MustPutProperty(monthlyDOWTriggerObj, "DaysOfWeek", monthlyDOWTrigger.DaysOfWeek)
			oleutil.MustPutProperty(monthlyDOWTriggerObj, "MonthsOfYear", monthlyDOWTrigger.MonthsOfYear)
			oleutil.MustPutProperty(monthlyDOWTriggerObj, "RandomDelay", monthlyDOWTrigger.RandomDelay)
			oleutil.MustPutProperty(monthlyDOWTriggerObj, "RunOnLastWeekOfMonth", monthlyDOWTrigger.RunOnLastWeekOfMonth)
			oleutil.MustPutProperty(monthlyDOWTriggerObj, "WeeksOfMonth", monthlyDOWTrigger.WeeksOfMonth)
		case TASK_TRIGGER_MONTHLY:
			monthlyTrigger := trigger.(MonthlyTrigger)
			monthlyTriggerObj := triggerObj.MustQueryInterface(ole.NewGUID("{97c45ef1-6b02-4a1a-9c0e-1ebfba1500ac}"))
			defer monthlyTriggerObj.Release()

			oleutil.MustPutProperty(monthlyTriggerObj, "DaysOfMonth", monthlyTrigger.DaysOfMonth)
			oleutil.MustPutProperty(monthlyTriggerObj, "MonthsOfYear", monthlyTrigger.MonthsOfYear)
			oleutil.MustPutProperty(monthlyTriggerObj, "RandomDelay", monthlyTrigger.RandomDelay)
			oleutil.MustPutProperty(monthlyTriggerObj, "RunOnLastDayOfMonth", monthlyTrigger.RunOnLastWeekOfMonth)
		case TASK_TRIGGER_REGISTRATION:
			registrationTrigger := trigger.(RegistrationTrigger)
			registrationTriggerObj := triggerObj.MustQueryInterface(ole.NewGUID("{4c8fec3a-c218-4e0c-b23d-629024db91a2}"))
			defer registrationTriggerObj.Release()

			oleutil.MustPutProperty(registrationTriggerObj, "Delay", registrationTrigger.Delay)
		case TASK_TRIGGER_TIME:
			timeTrigger := trigger.(TimeTrigger)
			timeTriggerObj := triggerObj.MustQueryInterface(ole.NewGUID("{b45747e0-eba7-4276-9f29-85c5bb300006}"))
			defer timeTriggerObj.Release()

			oleutil.MustPutProperty(timeTriggerObj, "RandomDelay", timeTrigger.RandomDelay)
		case TASK_TRIGGER_WEEKLY:
			weeklyTrigger := trigger.(WeeklyTrigger)
			weeklyTriggerObj := triggerObj.MustQueryInterface(ole.NewGUID("{5038fc98-82ff-436d-8728-a512a57c9dc1}"))
			defer weeklyTriggerObj.Release()

			oleutil.MustPutProperty(weeklyTriggerObj, "DaysOfWeek", weeklyTrigger.DaysOfWeek)
			oleutil.MustPutProperty(weeklyTriggerObj, "RandomDelay", weeklyTrigger.RandomDelay)
			oleutil.MustPutProperty(weeklyTriggerObj, "WeeksInterval", weeklyTrigger.WeeksInterval)
		case TASK_TRIGGER_SESSION_STATE_CHANGE:
			sessionStateChangeTrigger := trigger.(SessionStateChangeTrigger)
			sessionStateChangeTriggerObj := triggerObj.MustQueryInterface(ole.NewGUID("{754da71b-4385-4475-9dd9-598294fa3641}"))
			defer sessionStateChangeTriggerObj.Release()

			oleutil.MustPutProperty(sessionStateChangeTriggerObj, "Delay", sessionStateChangeTrigger.Delay)
			oleutil.MustPutProperty(sessionStateChangeTriggerObj, "StateChange", sessionStateChangeTrigger.StateChange)
			oleutil.MustPutProperty(sessionStateChangeTriggerObj, "UserId", sessionStateChangeTrigger.UserId)
			// need to find GUID
			/*case TASK_TRIGGER_CUSTOM_TRIGGER_01:
			return nil*/
		}
	}

	return nil
}

func checkTriggerType(triggerType int) bool {
	switch triggerType {
	case TASK_TRIGGER_BOOT:
		fallthrough
	case TASK_TRIGGER_DAILY:
		fallthrough
	case TASK_TRIGGER_EVENT:
		fallthrough
	case TASK_TRIGGER_IDLE:
		fallthrough
	case TASK_TRIGGER_LOGON:
		fallthrough
	case TASK_TRIGGER_MONTHLYDOW:
		fallthrough
	case TASK_TRIGGER_MONTHLY:
		fallthrough
	case TASK_TRIGGER_REGISTRATION:
		fallthrough
	case TASK_TRIGGER_TIME:
		fallthrough
	case TASK_TRIGGER_WEEKLY:
		fallthrough
	case TASK_TRIGGER_SESSION_STATE_CHANGE:
		return true
		// need to find GUID
		/*case TASK_TRIGGER_CUSTOM_TRIGGER_01:
		return nil*/
	default:
		return false
	}
}

func (r *RunningTask) Stop() error {
	stopResult := oleutil.MustCallMethod(r.taskObj, "Stop").Val
	if stopResult != 0 {
		return errors.New("cannot stop running task; access is denied")
	}

	r.taskObj.Release()

	return nil
}

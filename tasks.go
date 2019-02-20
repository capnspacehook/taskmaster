package taskmaster

import (
	"errors"
	"github.com/go-ole/go-ole"
	"time"

	"github.com/go-ole/go-ole/oleutil"
)

// AddExecAction adds an execute action to the task definition. The args
// parameter can have up to 32 $(ArgX) values, such as '/c $(Arg0) $(Arg1)'.
// This will allow the arguments to be dynamically entered when the task is run.
func (d *Definition) AddExecAction(path, args, workingDir, id string) {
	d.Actions = append(d.Actions, ExecAction{
		Path:       path,
		Args:       args,
		WorkingDir: workingDir,
		TaskAction: TaskAction{
			ID: id,
			taskActionTypeHolder: taskActionTypeHolder{
				actionType: TASK_ACTION_EXEC,
			},
		},
	})
}

// AddComHandlerAction adds a COM handler action to the task definition. The clisd
// parameter is the CLSID of the COM object that will get instantiated when the action
// executes, and the data parameter is the arguments passed to the COM object.
func (d *Definition) AddComHandlerAction(clsid, data, id string) {
	d.Actions = append(d.Actions, ComHandlerAction{
		ClassID: clsid,
		Data:    data,
		TaskAction: TaskAction{
			ID: id,
			taskActionTypeHolder: taskActionTypeHolder{
				actionType: TASK_ACTION_COM_HANDLER,
			},
		},
	})
}

func (d *Definition) AddBootTrigger(delay, id string, startBoundary, endBoundary time.Time, timeLimit, repetitionDuration, repetitionInterval string, stopAtDurationEnd, enabled bool) {
	startBoundaryStr := TimeToTaskDate(startBoundary)
	endBoundaryStr := TimeToTaskDate(endBoundary)

	d.Triggers = append(d.Triggers, BootTrigger{
		Delay: delay,
		TaskTrigger: TaskTrigger{
			Enabled:            enabled,
			EndBoundary:        endBoundaryStr,
			ExecutionTimeLimit: timeLimit,
			ID:                 id,
			RepetitionPattern: RepetitionPattern{
				RepetitionDuration: repetitionDuration,
				RepetitionInterval: repetitionInterval,
				StopAtDurationEnd:  stopAtDurationEnd,
			},
			StartBoundary: startBoundaryStr,
			taskTriggerTypeHolder: taskTriggerTypeHolder{
				triggerType: TASK_TRIGGER_BOOT,
			},
		},
	})
}

func (d *Definition) AddDailyTrigger(dayInterval DayInterval, randomDelay, id string, startBoundary, endBoundary time.Time, timeLimit, repetitionDuration, repetitionInterval string, stopAtDurationEnd, enabled bool) {
	startBoundaryStr := TimeToTaskDate(startBoundary)
	endBoundaryStr := TimeToTaskDate(endBoundary)

	d.Triggers = append(d.Triggers, DailyTrigger{
		DayInterval: dayInterval,
		RandomDelay: randomDelay,
		TaskTrigger: TaskTrigger{
			Enabled:            enabled,
			EndBoundary:        endBoundaryStr,
			ExecutionTimeLimit: timeLimit,
			ID:                 id,
			RepetitionPattern: RepetitionPattern{
				RepetitionDuration: repetitionDuration,
				RepetitionInterval: repetitionInterval,
				StopAtDurationEnd:  stopAtDurationEnd,
			},
			StartBoundary: startBoundaryStr,
			taskTriggerTypeHolder: taskTriggerTypeHolder{
				triggerType: TASK_TRIGGER_DAILY,
			},
		},
	})
}

func (d *Definition) AddEventTrigger(delay, subscription string, valueQueries map[string]string, id string, startBoundary, endBoundary time.Time, timeLimit, repetitionDuration, repetitionInterval string, stopAtDurationEnd, enabled bool) {
	startBoundaryStr := TimeToTaskDate(startBoundary)
	endBoundaryStr := TimeToTaskDate(endBoundary)

	d.Triggers = append(d.Triggers, EventTrigger{
		Delay:        delay,
		Subscription: subscription,
		ValueQueries: valueQueries,
		TaskTrigger: TaskTrigger{
			Enabled:            enabled,
			EndBoundary:        endBoundaryStr,
			ExecutionTimeLimit: timeLimit,
			ID:                 id,
			RepetitionPattern: RepetitionPattern{
				RepetitionDuration: repetitionDuration,
				RepetitionInterval: repetitionInterval,
				StopAtDurationEnd:  stopAtDurationEnd,
			},
			StartBoundary: startBoundaryStr,
			taskTriggerTypeHolder: taskTriggerTypeHolder{
				triggerType: TASK_TRIGGER_EVENT,
			},
		},
	})
}

func (d *Definition) AddIdleTrigger(id string, startBoundary, endBoundary time.Time, timeLimit, repetitionDuration, repetitionInterval string, stopAtDurationEnd, enabled bool) {
	startBoundaryStr := TimeToTaskDate(startBoundary)
	endBoundaryStr := TimeToTaskDate(endBoundary)

	d.Triggers = append(d.Triggers, IdleTrigger{
		TaskTrigger: TaskTrigger{
			Enabled:            enabled,
			EndBoundary:        endBoundaryStr,
			ExecutionTimeLimit: timeLimit,
			ID:                 id,
			RepetitionPattern: RepetitionPattern{
				RepetitionDuration: repetitionDuration,
				RepetitionInterval: repetitionInterval,
				StopAtDurationEnd:  stopAtDurationEnd,
			},
			StartBoundary: startBoundaryStr,
			taskTriggerTypeHolder: taskTriggerTypeHolder{
				triggerType: TASK_TRIGGER_IDLE,
			},
		},
	})
}

func (d *Definition) AddLogonTrigger(delay, userID, id string, startBoundary, endBoundary time.Time, timeLimit, repetitionDuration, repetitionInterval string, stopAtDurationEnd, enabled bool) {
	startBoundaryStr := TimeToTaskDate(startBoundary)
	endBoundaryStr := TimeToTaskDate(endBoundary)

	d.Triggers = append(d.Triggers, LogonTrigger{
		Delay:  delay,
		UserID: userID,
		TaskTrigger: TaskTrigger{
			Enabled:            enabled,
			EndBoundary:        endBoundaryStr,
			ExecutionTimeLimit: timeLimit,
			ID:                 id,
			RepetitionPattern: RepetitionPattern{
				RepetitionDuration: repetitionDuration,
				RepetitionInterval: repetitionInterval,
				StopAtDurationEnd:  stopAtDurationEnd,
			},
			StartBoundary: startBoundaryStr,
			taskTriggerTypeHolder: taskTriggerTypeHolder{
				triggerType: TASK_TRIGGER_LOGON,
			},
		},
	})
}

func (d *Definition) AddMonthlyDOWTrigger(dayOfWeek Day, weekOfMonth Week, monthOfYear Month, runOnLastWeekOfMonth bool, randomDelay, id string, startBoundary, endBoundary time.Time, timeLimit, repetitionDuration, repetitionInterval string, stopAtDurationEnd, enabled bool) {
	startBoundaryStr := TimeToTaskDate(startBoundary)
	endBoundaryStr := TimeToTaskDate(endBoundary)

	d.Triggers = append(d.Triggers, MonthlyDOWTrigger{
		DaysOfWeek:           dayOfWeek,
		MonthsOfYear:         monthOfYear,
		RandomDelay:          randomDelay,
		RunOnLastWeekOfMonth: runOnLastWeekOfMonth,
		WeeksOfMonth:         weekOfMonth,
		TaskTrigger: TaskTrigger{
			Enabled:            enabled,
			EndBoundary:        endBoundaryStr,
			ExecutionTimeLimit: timeLimit,
			ID:                 id,
			RepetitionPattern: RepetitionPattern{
				RepetitionDuration: repetitionDuration,
				RepetitionInterval: repetitionInterval,
				StopAtDurationEnd:  stopAtDurationEnd,
			},
			StartBoundary: startBoundaryStr,
			taskTriggerTypeHolder: taskTriggerTypeHolder{
				triggerType: TASK_TRIGGER_MONTHLYDOW,
			},
		},
	})
}

func (d *Definition) AddMonthlyTrigger(dayOfMonth int, monthOfYear Month, randomDelay, id string, startBoundary, endBoundary time.Time, timeLimit, repetitionDuration, repetitionInterval string, stopAtDurationEnd, enabled bool) error {
	monthDay, err := IntToDayOfMonth(dayOfMonth)
	if err != nil {
		return err
	}
	startBoundaryStr := TimeToTaskDate(startBoundary)
	endBoundaryStr := TimeToTaskDate(endBoundary)

	d.Triggers = append(d.Triggers, MonthlyTrigger{
		DaysOfMonth:  monthDay,
		MonthsOfYear: monthOfYear,
		RandomDelay:  randomDelay,
		TaskTrigger: TaskTrigger{
			Enabled:            enabled,
			EndBoundary:        endBoundaryStr,
			ExecutionTimeLimit: timeLimit,
			ID:                 id,
			RepetitionPattern: RepetitionPattern{
				RepetitionDuration: repetitionDuration,
				RepetitionInterval: repetitionInterval,
				StopAtDurationEnd:  stopAtDurationEnd,
			},
			StartBoundary: startBoundaryStr,
			taskTriggerTypeHolder: taskTriggerTypeHolder{
				triggerType: TASK_TRIGGER_MONTHLY,
			},
		},
	})

	return nil
}

func (d *Definition) AddRegistrationTrigger(delay, id string, startBoundary, endBoundary time.Time, timeLimit, repetitionDuration, repetitionInterval string, stopAtDurationEnd, enabled bool) {
	startBoundaryStr := TimeToTaskDate(startBoundary)
	endBoundaryStr := TimeToTaskDate(endBoundary)

	d.Triggers = append(d.Triggers, RegistrationTrigger{
		Delay: delay,
		TaskTrigger: TaskTrigger{
			Enabled:            enabled,
			EndBoundary:        endBoundaryStr,
			ExecutionTimeLimit: timeLimit,
			ID:                 id,
			RepetitionPattern: RepetitionPattern{
				RepetitionDuration: repetitionDuration,
				RepetitionInterval: repetitionInterval,
				StopAtDurationEnd:  stopAtDurationEnd,
			},
			StartBoundary: startBoundaryStr,
			taskTriggerTypeHolder: taskTriggerTypeHolder{
				triggerType: TASK_TRIGGER_REGISTRATION,
			},
		},
	})
}

func (d *Definition) AddSessionStateChangeTrigger(userID string, stateChange TaskSessionStateChangeType, delay, id string, startBoundary, endBoundary time.Time, timeLimit, repetitionDuration, repetitionInterval string, stopAtDurationEnd, enabled bool) {
	startBoundaryStr := TimeToTaskDate(startBoundary)
	endBoundaryStr := TimeToTaskDate(endBoundary)

	d.Triggers = append(d.Triggers, SessionStateChangeTrigger{
		Delay:       delay,
		StateChange: stateChange,
		UserId:      userID,
		TaskTrigger: TaskTrigger{
			Enabled:            enabled,
			EndBoundary:        endBoundaryStr,
			ExecutionTimeLimit: timeLimit,
			ID:                 id,
			RepetitionPattern: RepetitionPattern{
				RepetitionDuration: repetitionDuration,
				RepetitionInterval: repetitionInterval,
				StopAtDurationEnd:  stopAtDurationEnd,
			},
			StartBoundary: startBoundaryStr,
			taskTriggerTypeHolder: taskTriggerTypeHolder{
				triggerType: TASK_TRIGGER_SESSION_STATE_CHANGE,
			},
		},
	})
}

func (d *Definition) AddTimeTrigger(randomDelay, id string, startBoundary, endBoundary time.Time, timeLimit, repetitionDuration, repetitionInterval string, stopAtDurationEnd, enabled bool) {
	startBoundaryStr := TimeToTaskDate(startBoundary)
	endBoundaryStr := TimeToTaskDate(endBoundary)

	d.Triggers = append(d.Triggers, TimeTrigger{
		RandomDelay: randomDelay,
		TaskTrigger: TaskTrigger{
			Enabled:            enabled,
			EndBoundary:        endBoundaryStr,
			ExecutionTimeLimit: timeLimit,
			ID:                 id,
			RepetitionPattern: RepetitionPattern{
				RepetitionDuration: repetitionDuration,
				RepetitionInterval: repetitionInterval,
				StopAtDurationEnd:  stopAtDurationEnd,
			},
			StartBoundary: startBoundaryStr,
			taskTriggerTypeHolder: taskTriggerTypeHolder{
				triggerType: TASK_TRIGGER_TIME,
			},
		},
	})
}

func (d *Definition) AddWeeklyTrigger(dayOfWeek Day, weekInterval WeekInterval, randomDelay, id string, startBoundary, endBoundary time.Time, timeLimit, repetitionDuration, repetitionInterval string, stopAtDurationEnd, enabled bool) {
	startBoundaryStr := TimeToTaskDate(startBoundary)
	endBoundaryStr := TimeToTaskDate(endBoundary)

	d.Triggers = append(d.Triggers, WeeklyTrigger{
		DaysOfWeek:   dayOfWeek,
		RandomDelay:  randomDelay,
		WeekInterval: weekInterval,
		TaskTrigger: TaskTrigger{
			Enabled:            enabled,
			EndBoundary:        endBoundaryStr,
			ExecutionTimeLimit: timeLimit,
			ID:                 id,
			RepetitionPattern: RepetitionPattern{
				RepetitionDuration: repetitionDuration,
				RepetitionInterval: repetitionInterval,
				StopAtDurationEnd:  stopAtDurationEnd,
			},
			StartBoundary: startBoundaryStr,
			taskTriggerTypeHolder: taskTriggerTypeHolder{
				triggerType: TASK_TRIGGER_WEEKLY,
			},
		},
	})
}

// Refresh refreshes all of the local instance variables of the running task.
// https://docs.microsoft.com/en-us/windows/desktop/api/taskschd/nf-taskschd-irunningtask-refresh
func (r RunningTask) Refresh() error {
	_, err := oleutil.CallMethod(r.taskObj, "Refresh")
	if err != nil {
		return err
	}

	return nil
}

// Stop kills a running task.
// https://docs.microsoft.com/en-us/windows/desktop/api/taskschd/nf-taskschd-irunningtask-stop
func (r *RunningTask) Stop() error {
	_, err := oleutil.CallMethod(r.taskObj, "Stop")
	if err != nil {
		return errors.New("cannot stop running task; access is denied")
	}

	r.taskObj.Release()

	return nil
}

// Release frees the running task COM object. Must be called before
// program termination to avoid memory leaks.
func (r *RunningTask) Release() {
	if !r.isReleased {
		r.taskObj.Release()
		r.isReleased = true
	}
}

// Run starts an instance of a registered task. If the task was started successfully,
// a pointer to a running task will be returned.
// https://docs.microsoft.com/en-us/windows/desktop/api/taskschd/nf-taskschd-iregisteredtask-runex
func (r *RegisteredTask) Run(args []string, flags TaskRunFlags, sessionID int, user string) (*RunningTask, error) {
	if !r.Enabled {
		return nil, errors.New("cannot run a disabled task")
	}

	runningTaskObj, err := oleutil.CallMethod(r.taskObj, "RunEx", args, int(flags), sessionID, user)
	if err != nil {
		return nil, err
	}

	runningTask := parseRunningTask(runningTaskObj.ToIDispatch())

	return &runningTask, nil
}

// GetInstances returns all of the currently running instances of a registered task.
// https://docs.microsoft.com/en-us/windows/desktop/api/taskschd/nf-taskschd-iregisteredtask-getinstances
func (r *RegisteredTask) GetInstances() ([]*RunningTask, error) {
	runningTasks, err := oleutil.CallMethod(r.taskObj, "GetInstances", 0)
	if err != nil {
		return nil, err
	}

	runningTasksObj := runningTasks.ToIDispatch()
	defer runningTasksObj.Release()
	var parsedRunningTasks []*RunningTask

	oleutil.ForEach(runningTasksObj, func(v *ole.VARIANT) error {
		runningTaskObj := v.ToIDispatch()

		parsedRunningTask := parseRunningTask(runningTaskObj)
		parsedRunningTasks = append(parsedRunningTasks, &parsedRunningTask)

		return nil
	})

	return parsedRunningTasks, nil
}

// Stop kills all running instances of the registered task that the current
// user has access to. If all instances were killed, Stop returns true,
// otherwise Stop returns false.
// https://docs.microsoft.com/en-us/windows/desktop/api/taskschd/nf-taskschd-iregisteredtask-stop
func (r *RegisteredTask) Stop() bool {
	ret, _ := oleutil.CallMethod(r.taskObj, "Stop", 0)
	if ret.Val != 0 {
		return false
	}

	return true
}

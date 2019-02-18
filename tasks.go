package taskmaster

import (
	"errors"
	"time"

	"github.com/go-ole/go-ole/oleutil"
)

// AddExecAction adds an execute action to the task definition. The args
// parameter can have up to 32 $(ArgX) values, such as '/c $(Arg0) $(Arg1)'.
// This will allow the arguments to be dynamically entered when the task is run
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
// executes, and the data parameter is the arguments passed to the COM object
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

func (d *Definition) AddBootTrigger(delay, id string, startBoundary, endBoundary time.Time, timeLimit, repitionDuration, repitionInterval string, stopAtDurationEnd, enabled bool) {
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
				RepitionDuration:  repitionDuration,
				RepitionInterval:  repitionInterval,
				StopAtDurationEnd: stopAtDurationEnd,
			},
			StartBoundary: startBoundaryStr,
			taskTriggerTypeHolder: taskTriggerTypeHolder{
				triggerType: TASK_TRIGGER_BOOT,
			},
		},
	})
}

func (d *Definition) AddDailyTrigger(dayInterval DayInterval, randomDelay, id string, startBoundary, endBoundary time.Time, timeLimit, repitionDuration, repitionInterval string, stopAtDurationEnd, enabled bool) {
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
				RepitionDuration:  repitionDuration,
				RepitionInterval:  repitionInterval,
				StopAtDurationEnd: stopAtDurationEnd,
			},
			StartBoundary: startBoundaryStr,
			taskTriggerTypeHolder: taskTriggerTypeHolder{
				triggerType: TASK_TRIGGER_DAILY,
			},
		},
	})
}

func (d *Definition) AddEventTrigger(delay, subscription string, valueQueries map[string]string, id string, startBoundary, endBoundary time.Time, timeLimit, repitionDuration, repitionInterval string, stopAtDurationEnd, enabled bool) {
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
				RepitionDuration:  repitionDuration,
				RepitionInterval:  repitionInterval,
				StopAtDurationEnd: stopAtDurationEnd,
			},
			StartBoundary: startBoundaryStr,
			taskTriggerTypeHolder: taskTriggerTypeHolder{
				triggerType: TASK_TRIGGER_EVENT,
			},
		},
	})
}

func (d *Definition) AddIdleTrigger(id string, startBoundary, endBoundary time.Time, timeLimit, repitionDuration, repitionInterval string, stopAtDurationEnd, enabled bool) {
	startBoundaryStr := TimeToTaskDate(startBoundary)
	endBoundaryStr := TimeToTaskDate(endBoundary)

	d.Triggers = append(d.Triggers, IdleTrigger{
		TaskTrigger: TaskTrigger{
			Enabled:            enabled,
			EndBoundary:        endBoundaryStr,
			ExecutionTimeLimit: timeLimit,
			ID:                 id,
			RepetitionPattern: RepetitionPattern{
				RepitionDuration:  repitionDuration,
				RepitionInterval:  repitionInterval,
				StopAtDurationEnd: stopAtDurationEnd,
			},
			StartBoundary: startBoundaryStr,
			taskTriggerTypeHolder: taskTriggerTypeHolder{
				triggerType: TASK_TRIGGER_IDLE,
			},
		},
	})
}

func (d *Definition) AddLogonTrigger(delay, userID, id string, startBoundary, endBoundary time.Time, timeLimit, repitionDuration, repitionInterval string, stopAtDurationEnd, enabled bool) {
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
				RepitionDuration:  repitionDuration,
				RepitionInterval:  repitionInterval,
				StopAtDurationEnd: stopAtDurationEnd,
			},
			StartBoundary: startBoundaryStr,
			taskTriggerTypeHolder: taskTriggerTypeHolder{
				triggerType: TASK_TRIGGER_LOGON,
			},
		},
	})
}

func (d *Definition) AddMonthlyDOWTrigger(dayOfWeek Day, weekOfMonth Week, monthOfYear Month, runOnLastWeekOfMonth bool, randomDelay, id string, startBoundary, endBoundary time.Time, timeLimit, repitionDuration, repitionInterval string, stopAtDurationEnd, enabled bool) {
	startBoundaryStr := TimeToTaskDate(startBoundary)
	endBoundaryStr := TimeToTaskDate(endBoundary)

	d.Triggers = append(d.Triggers, MonthlyDOWTrigger{
		DayOfWeek:            dayOfWeek,
		MonthOfYear:          monthOfYear,
		RandomDelay:          randomDelay,
		RunOnLastWeekOfMonth: runOnLastWeekOfMonth,
		WeekOfMonth:          weekOfMonth,
		TaskTrigger: TaskTrigger{
			Enabled:            enabled,
			EndBoundary:        endBoundaryStr,
			ExecutionTimeLimit: timeLimit,
			ID:                 id,
			RepetitionPattern: RepetitionPattern{
				RepitionDuration:  repitionDuration,
				RepitionInterval:  repitionInterval,
				StopAtDurationEnd: stopAtDurationEnd,
			},
			StartBoundary: startBoundaryStr,
			taskTriggerTypeHolder: taskTriggerTypeHolder{
				triggerType: TASK_TRIGGER_MONTHLYDOW,
			},
		},
	})
}

func (d *Definition) AddMonthlyTrigger(dayOfMonth int, monthOfYear Month, randomDelay, id string, startBoundary, endBoundary time.Time, timeLimit, repitionDuration, repitionInterval string, stopAtDurationEnd, enabled bool) error {
	monthDay, err := IntToDayOfMonth(dayOfMonth)
	if err != nil {
		return err
	}
	startBoundaryStr := TimeToTaskDate(startBoundary)
	endBoundaryStr := TimeToTaskDate(endBoundary)

	d.Triggers = append(d.Triggers, MonthlyTrigger{
		DayOfMonth:  monthDay,
		MonthOfYear: monthOfYear,
		RandomDelay: randomDelay,
		TaskTrigger: TaskTrigger{
			Enabled:            enabled,
			EndBoundary:        endBoundaryStr,
			ExecutionTimeLimit: timeLimit,
			ID:                 id,
			RepetitionPattern: RepetitionPattern{
				RepitionDuration:  repitionDuration,
				RepitionInterval:  repitionInterval,
				StopAtDurationEnd: stopAtDurationEnd,
			},
			StartBoundary: startBoundaryStr,
			taskTriggerTypeHolder: taskTriggerTypeHolder{
				triggerType: TASK_TRIGGER_MONTHLY,
			},
		},
	})

	return nil
}

func (d *Definition) AddRegistrationTrigger(delay, id string, startBoundary, endBoundary time.Time, timeLimit, repitionDuration, repitionInterval string, stopAtDurationEnd, enabled bool) {
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
				RepitionDuration:  repitionDuration,
				RepitionInterval:  repitionInterval,
				StopAtDurationEnd: stopAtDurationEnd,
			},
			StartBoundary: startBoundaryStr,
			taskTriggerTypeHolder: taskTriggerTypeHolder{
				triggerType: TASK_TRIGGER_REGISTRATION,
			},
		},
	})
}

func (d *Definition) AddSessionStateChangeTrigger(userID string, stateChange TaskSessionStateChangeType, delay, id string, startBoundary, endBoundary time.Time, timeLimit, repitionDuration, repitionInterval string, stopAtDurationEnd, enabled bool) {
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
				RepitionDuration:  repitionDuration,
				RepitionInterval:  repitionInterval,
				StopAtDurationEnd: stopAtDurationEnd,
			},
			StartBoundary: startBoundaryStr,
			taskTriggerTypeHolder: taskTriggerTypeHolder{
				triggerType: TASK_TRIGGER_SESSION_STATE_CHANGE,
			},
		},
	})
}

func (d *Definition) AddTimeTrigger(randomDelay, id string, startBoundary, endBoundary time.Time, timeLimit, repitionDuration, repitionInterval string, stopAtDurationEnd, enabled bool) {
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
				RepitionDuration:  repitionDuration,
				RepitionInterval:  repitionInterval,
				StopAtDurationEnd: stopAtDurationEnd,
			},
			StartBoundary: startBoundaryStr,
			taskTriggerTypeHolder: taskTriggerTypeHolder{
				triggerType: TASK_TRIGGER_TIME,
			},
		},
	})
}

func (d *Definition) AddWeeklyTrigger(dayOfWeek Day, weekInterval WeekInterval, randomDelay, id string, startBoundary, endBoundary time.Time, timeLimit, repitionDuration, repitionInterval string, stopAtDurationEnd, enabled bool) {
	startBoundaryStr := TimeToTaskDate(startBoundary)
	endBoundaryStr := TimeToTaskDate(endBoundary)

	d.Triggers = append(d.Triggers, WeeklyTrigger{
		DayOfWeek:    dayOfWeek,
		RandomDelay:  randomDelay,
		WeekInterval: weekInterval,
		TaskTrigger: TaskTrigger{
			Enabled:            enabled,
			EndBoundary:        endBoundaryStr,
			ExecutionTimeLimit: timeLimit,
			ID:                 id,
			RepetitionPattern: RepetitionPattern{
				RepitionDuration:  repitionDuration,
				RepitionInterval:  repitionInterval,
				StopAtDurationEnd: stopAtDurationEnd,
			},
			StartBoundary: startBoundaryStr,
			taskTriggerTypeHolder: taskTriggerTypeHolder{
				triggerType: TASK_TRIGGER_WEEKLY,
			},
		},
	})
}

// Stop kills a running task
func (r *RunningTask) Stop() error {
	stopResult := oleutil.MustCallMethod(r.taskObj, "Stop").Val
	if stopResult != 0 {
		return errors.New("cannot stop running task; access is denied")
	}

	r.taskObj.Release()

	return nil
}

// Release frees the running task COM object. Must be called before
// program termination to avoid memory leaks
func (r *RunningTask) Release() {
	r.taskObj.Release()
}

// Run starts an instance of a registered task. If the task was started successfully,
// a pointer to a running task will be returned
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

// Stop kills all running instances of the registered task that the current
// user has access to. If all instances were killed, Stop returns true,
// otherwise Stop returns false
func (r *RegisteredTask) Stop() bool {
	ret, _ := oleutil.CallMethod(r.taskObj, "Stop", 0)
	if ret.Val != 0 {
		return false
	}

	return true
}

// Release frees the registered task COM object. Must be called before
// program termination to avoid memory leaks
func (r *RegisteredTask) Release() {
	r.taskObj.Release()
}

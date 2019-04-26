// +build windows

package taskmaster

import (
	"fmt"
	"time"

	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
	"github.com/rickb777/date/period"
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

// AddMessageBoxAction adds a MessageBox action to the task definition. The title
// parameter is the title of the MessageBox window, and the message parameter is the
// message the MessageBox will display. Be aware this action type is not supported on
// modern versions of Windows.
func (d *Definition) AddMessageBoxAction(title, message, id string) {
	d.Actions = append(d.Actions, MessageBoxAction{
		Title:   title,
		Message: message,
		TaskAction: TaskAction{
			ID: id,
			taskActionTypeHolder: taskActionTypeHolder{
				actionType: TASK_ACTION_SHOW_MESSAGE,
			},
		},
	})
}

func (d *Definition) AddBootTrigger(delay period.Period) {
	d.AddBootTriggerEx(delay, "", time.Time{}, time.Time{}, period.Period{}, period.Period{}, period.Period{}, false, true)
}

func (d *Definition) AddBootTriggerEx(delay period.Period, id string, startBoundary, endBoundary time.Time, timeLimit, repetitionDuration, repetitionInterval period.Period, stopAtDurationEnd, enabled bool) {
	d.Triggers = append(d.Triggers, BootTrigger{
		Delay: delay,
		TaskTrigger: TaskTrigger{
			Enabled:            enabled,
			EndBoundary:        endBoundary,
			ExecutionTimeLimit: timeLimit,
			ID:                 id,
			RepetitionPattern: RepetitionPattern{
				RepetitionDuration: repetitionDuration,
				RepetitionInterval: repetitionInterval,
				StopAtDurationEnd:  stopAtDurationEnd,
			},
			StartBoundary: startBoundary,
			taskTriggerTypeHolder: taskTriggerTypeHolder{
				triggerType: TASK_TRIGGER_BOOT,
			},
		},
	})
}

func (d *Definition) AddDailyTrigger(dayInterval DayInterval, randomDelay period.Period, startBoundary time.Time) {
	d.AddDailyTriggerEx(dayInterval, randomDelay, "", startBoundary, time.Time{}, period.Period{}, period.Period{}, period.Period{}, false, true)
}

func (d *Definition) AddDailyTriggerEx(dayInterval DayInterval, randomDelay period.Period, id string, startBoundary, endBoundary time.Time, timeLimit, repetitionDuration, repetitionInterval period.Period, stopAtDurationEnd, enabled bool) {
	d.Triggers = append(d.Triggers, DailyTrigger{
		DayInterval: dayInterval,
		RandomDelay: randomDelay,
		TaskTrigger: TaskTrigger{
			Enabled:            enabled,
			EndBoundary:        endBoundary,
			ExecutionTimeLimit: timeLimit,
			ID:                 id,
			RepetitionPattern: RepetitionPattern{
				RepetitionDuration: repetitionDuration,
				RepetitionInterval: repetitionInterval,
				StopAtDurationEnd:  stopAtDurationEnd,
			},
			StartBoundary: startBoundary,
			taskTriggerTypeHolder: taskTriggerTypeHolder{
				triggerType: TASK_TRIGGER_DAILY,
			},
		},
	})
}

func (d *Definition) AddEventTrigger(delay period.Period, subscription string, valueQueries map[string]string) {
	d.AddEventTriggerEx(delay, subscription, valueQueries, "", time.Time{}, time.Time{}, period.Period{}, period.Period{}, period.Period{}, false, true)
}

func (d *Definition) AddEventTriggerEx(delay period.Period, subscription string, valueQueries map[string]string, id string, startBoundary, endBoundary time.Time, timeLimit, repetitionDuration, repetitionInterval period.Period, stopAtDurationEnd, enabled bool) {
	d.Triggers = append(d.Triggers, EventTrigger{
		Delay:        delay,
		Subscription: subscription,
		ValueQueries: valueQueries,
		TaskTrigger: TaskTrigger{
			Enabled:            enabled,
			EndBoundary:        endBoundary,
			ExecutionTimeLimit: timeLimit,
			ID:                 id,
			RepetitionPattern: RepetitionPattern{
				RepetitionDuration: repetitionDuration,
				RepetitionInterval: repetitionInterval,
				StopAtDurationEnd:  stopAtDurationEnd,
			},
			StartBoundary: startBoundary,
			taskTriggerTypeHolder: taskTriggerTypeHolder{
				triggerType: TASK_TRIGGER_EVENT,
			},
		},
	})
}

func (d *Definition) AddIdleTrigger() {
	d.AddIdleTriggerEx("", time.Time{}, time.Time{}, period.Period{}, period.Period{}, period.Period{}, false, true)
}

func (d *Definition) AddIdleTriggerEx(id string, startBoundary, endBoundary time.Time, timeLimit, repetitionDuration, repetitionInterval period.Period, stopAtDurationEnd, enabled bool) {
	d.Triggers = append(d.Triggers, IdleTrigger{
		TaskTrigger: TaskTrigger{
			Enabled:            enabled,
			EndBoundary:        endBoundary,
			ExecutionTimeLimit: timeLimit,
			ID:                 id,
			RepetitionPattern: RepetitionPattern{
				RepetitionDuration: repetitionDuration,
				RepetitionInterval: repetitionInterval,
				StopAtDurationEnd:  stopAtDurationEnd,
			},
			StartBoundary: startBoundary,
			taskTriggerTypeHolder: taskTriggerTypeHolder{
				triggerType: TASK_TRIGGER_IDLE,
			},
		},
	})
}

func (d *Definition) AddLogonTrigger(delay period.Period, userID string) {
	d.AddLogonTriggerEx(delay, userID, "", time.Time{}, time.Time{}, period.Period{}, period.Period{}, period.Period{}, false, true)
}

func (d *Definition) AddLogonTriggerEx(delay period.Period, userID, id string, startBoundary, endBoundary time.Time, timeLimit, repetitionDuration, repetitionInterval period.Period, stopAtDurationEnd, enabled bool) {
	d.Triggers = append(d.Triggers, LogonTrigger{
		Delay:  delay,
		UserID: userID,
		TaskTrigger: TaskTrigger{
			Enabled:            enabled,
			EndBoundary:        endBoundary,
			ExecutionTimeLimit: timeLimit,
			ID:                 id,
			RepetitionPattern: RepetitionPattern{
				RepetitionDuration: repetitionDuration,
				RepetitionInterval: repetitionInterval,
				StopAtDurationEnd:  stopAtDurationEnd,
			},
			StartBoundary: startBoundary,
			taskTriggerTypeHolder: taskTriggerTypeHolder{
				triggerType: TASK_TRIGGER_LOGON,
			},
		},
	})
}

func (d *Definition) AddMonthlyDOWTrigger(dayOfWeek Day, weekOfMonth Week, monthOfYear Month, runOnLastWeekOfMonth bool, randomDelay period.Period, startBoundary time.Time) {
	d.AddMonthlyDOWTriggerEx(dayOfWeek, weekOfMonth, monthOfYear, runOnLastWeekOfMonth, randomDelay, "", startBoundary, time.Time{}, period.Period{}, period.Period{}, period.Period{}, false, true)
}

func (d *Definition) AddMonthlyDOWTriggerEx(dayOfWeek Day, weekOfMonth Week, monthOfYear Month, runOnLastWeekOfMonth bool, randomDelay period.Period, id string, startBoundary, endBoundary time.Time, timeLimit, repetitionDuration, repetitionInterval period.Period, stopAtDurationEnd, enabled bool) {
	d.Triggers = append(d.Triggers, MonthlyDOWTrigger{
		DaysOfWeek:           dayOfWeek,
		MonthsOfYear:         monthOfYear,
		RandomDelay:          randomDelay,
		RunOnLastWeekOfMonth: runOnLastWeekOfMonth,
		WeeksOfMonth:         weekOfMonth,
		TaskTrigger: TaskTrigger{
			Enabled:            enabled,
			EndBoundary:        endBoundary,
			ExecutionTimeLimit: timeLimit,
			ID:                 id,
			RepetitionPattern: RepetitionPattern{
				RepetitionDuration: repetitionDuration,
				RepetitionInterval: repetitionInterval,
				StopAtDurationEnd:  stopAtDurationEnd,
			},
			StartBoundary: startBoundary,
			taskTriggerTypeHolder: taskTriggerTypeHolder{
				triggerType: TASK_TRIGGER_MONTHLYDOW,
			},
		},
	})
}

func (d *Definition) AddMonthlyTrigger(dayOfMonth int, monthOfYear Month, randomDelay period.Period, startBoundary time.Time) {
	d.AddMonthlyTriggerEx(dayOfMonth, monthOfYear, randomDelay, "", startBoundary, time.Time{}, period.Period{}, period.Period{}, period.Period{}, false, true)
}

func (d *Definition) AddMonthlyTriggerEx(dayOfMonth int, monthOfYear Month, randomDelay period.Period, id string, startBoundary, endBoundary time.Time, timeLimit, repetitionDuration, repetitionInterval period.Period, stopAtDurationEnd, enabled bool) error {
	monthDay, err := IntToDayOfMonth(dayOfMonth)
	if err != nil {
		return err
	}
	d.Triggers = append(d.Triggers, MonthlyTrigger{
		DaysOfMonth:  monthDay,
		MonthsOfYear: monthOfYear,
		RandomDelay:  randomDelay,
		TaskTrigger: TaskTrigger{
			Enabled:            enabled,
			EndBoundary:        endBoundary,
			ExecutionTimeLimit: timeLimit,
			ID:                 id,
			RepetitionPattern: RepetitionPattern{
				RepetitionDuration: repetitionDuration,
				RepetitionInterval: repetitionInterval,
				StopAtDurationEnd:  stopAtDurationEnd,
			},
			StartBoundary: startBoundary,
			taskTriggerTypeHolder: taskTriggerTypeHolder{
				triggerType: TASK_TRIGGER_MONTHLY,
			},
		},
	})

	return nil
}

func (d *Definition) AddRegistrationTrigger(delay period.Period) {
	d.AddRegistrationTriggerEx(delay, "", time.Time{}, time.Time{}, period.Period{}, period.Period{}, period.Period{}, false, true)
}

func (d *Definition) AddRegistrationTriggerEx(delay period.Period, id string, startBoundary, endBoundary time.Time, timeLimit, repetitionDuration, repetitionInterval period.Period, stopAtDurationEnd, enabled bool) {
	d.Triggers = append(d.Triggers, RegistrationTrigger{
		Delay: delay,
		TaskTrigger: TaskTrigger{
			Enabled:            enabled,
			EndBoundary:        endBoundary,
			ExecutionTimeLimit: timeLimit,
			ID:                 id,
			RepetitionPattern: RepetitionPattern{
				RepetitionDuration: repetitionDuration,
				RepetitionInterval: repetitionInterval,
				StopAtDurationEnd:  stopAtDurationEnd,
			},
			StartBoundary: startBoundary,
			taskTriggerTypeHolder: taskTriggerTypeHolder{
				triggerType: TASK_TRIGGER_REGISTRATION,
			},
		},
	})
}

func (d *Definition) AddSessionStateChangeTrigger(userID string, stateChange TaskSessionStateChangeType, delay period.Period) {
	d.AddSessionStateChangeTriggerEx(userID, stateChange, delay, "", time.Time{}, time.Time{}, period.Period{}, period.Period{}, period.Period{}, false, true)
}

func (d *Definition) AddSessionStateChangeTriggerEx(userID string, stateChange TaskSessionStateChangeType, delay period.Period, id string, startBoundary, endBoundary time.Time, timeLimit, repetitionDuration, repetitionInterval period.Period, stopAtDurationEnd, enabled bool) {
	d.Triggers = append(d.Triggers, SessionStateChangeTrigger{
		Delay:       delay,
		StateChange: stateChange,
		UserId:      userID,
		TaskTrigger: TaskTrigger{
			Enabled:            enabled,
			EndBoundary:        endBoundary,
			ExecutionTimeLimit: timeLimit,
			ID:                 id,
			RepetitionPattern: RepetitionPattern{
				RepetitionDuration: repetitionDuration,
				RepetitionInterval: repetitionInterval,
				StopAtDurationEnd:  stopAtDurationEnd,
			},
			StartBoundary: startBoundary,
			taskTriggerTypeHolder: taskTriggerTypeHolder{
				triggerType: TASK_TRIGGER_SESSION_STATE_CHANGE,
			},
		},
	})
}

func (d *Definition) AddTimeTrigger(randomDelay period.Period, startBoundary time.Time) {
	d.AddTimeTriggerEx(randomDelay, "", startBoundary, time.Time{}, period.Period{}, period.Period{}, period.Period{}, false, true)
}

func (d *Definition) AddTimeTriggerEx(randomDelay period.Period, id string, startBoundary, endBoundary time.Time, timeLimit, repetitionDuration, repetitionInterval period.Period, stopAtDurationEnd, enabled bool) {
	d.Triggers = append(d.Triggers, TimeTrigger{
		RandomDelay: randomDelay,
		TaskTrigger: TaskTrigger{
			Enabled:            enabled,
			EndBoundary:        endBoundary,
			ExecutionTimeLimit: timeLimit,
			ID:                 id,
			RepetitionPattern: RepetitionPattern{
				RepetitionDuration: repetitionDuration,
				RepetitionInterval: repetitionInterval,
				StopAtDurationEnd:  stopAtDurationEnd,
			},
			StartBoundary: startBoundary,
			taskTriggerTypeHolder: taskTriggerTypeHolder{
				triggerType: TASK_TRIGGER_TIME,
			},
		},
	})
}

func (d *Definition) AddWeeklyTrigger(dayOfWeek Day, weekInterval WeekInterval, randomDelay period.Period, startBoundary time.Time) {
	d.AddWeeklyTriggerEx(dayOfWeek, weekInterval, randomDelay, "", startBoundary, time.Time{}, period.Period{}, period.Period{}, period.Period{}, false, true)
}

func (d *Definition) AddWeeklyTriggerEx(dayOfWeek Day, weekInterval WeekInterval, randomDelay period.Period, id string, startBoundary, endBoundary time.Time, timeLimit, repetitionDuration, repetitionInterval period.Period, stopAtDurationEnd, enabled bool) {
	d.Triggers = append(d.Triggers, WeeklyTrigger{
		DaysOfWeek:   dayOfWeek,
		RandomDelay:  randomDelay,
		WeekInterval: weekInterval,
		TaskTrigger: TaskTrigger{
			Enabled:            enabled,
			EndBoundary:        endBoundary,
			ExecutionTimeLimit: timeLimit,
			ID:                 id,
			RepetitionPattern: RepetitionPattern{
				RepetitionDuration: repetitionDuration,
				RepetitionInterval: repetitionInterval,
				StopAtDurationEnd:  stopAtDurationEnd,
			},
			StartBoundary: startBoundary,
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
		return fmt.Errorf("error calling Refresh on %s IRunningTask: %s", r.Path, err)
	}

	return nil
}

// Stop kills a running task.
// https://docs.microsoft.com/en-us/windows/desktop/api/taskschd/nf-taskschd-irunningtask-stop
func (r *RunningTask) Stop() error {
	_, err := oleutil.CallMethod(r.taskObj, "Stop")
	if err != nil {
		return fmt.Errorf("error calling Stop on %s IRunningTask: %s", r.Path, err)
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

func (r *RegisteredTask) Run(args []string) (*RunningTask, error) {
	return r.RunEx(args, TASK_RUN_AS_SELF, 0, "")
}

// RunEx starts an instance of a registered task. If the task was started successfully,
// a pointer to a running task will be returned.
// https://docs.microsoft.com/en-us/windows/desktop/api/taskschd/nf-taskschd-iregisteredtask-runex
func (r *RegisteredTask) RunEx(args []string, flags TaskRunFlags, sessionID int, user string) (*RunningTask, error) {
	if !r.Enabled {
		return nil, fmt.Errorf("error calling RunEx on %s IRegisteredTask: cannot run a disabled task", r.Path)
	}

	runningTaskObj, err := oleutil.CallMethod(r.taskObj, "RunEx", args, int(flags), sessionID, user)
	if err != nil {
		return nil, fmt.Errorf("error calling RunEx on %s IRegisteredTask: %s", r.Path, err)
	}

	runningTask := parseRunningTask(runningTaskObj.ToIDispatch())

	return runningTask, nil
}

// GetInstances returns all of the currently running instances of a registered task.
// https://docs.microsoft.com/en-us/windows/desktop/api/taskschd/nf-taskschd-iregisteredtask-getinstances
func (r *RegisteredTask) GetInstances() ([]*RunningTask, error) {
	runningTasks, err := oleutil.CallMethod(r.taskObj, "GetInstances", 0)
	if err != nil {
		return nil, fmt.Errorf("error calling RunEx on %s IRegisteredTask: %s", r.Path, err)
	}

	runningTasksObj := runningTasks.ToIDispatch()
	defer runningTasksObj.Release()
	var parsedRunningTasks []*RunningTask

	oleutil.ForEach(runningTasksObj, func(v *ole.VARIANT) error {
		runningTaskObj := v.ToIDispatch()

		parsedRunningTask := parseRunningTask(runningTaskObj)
		parsedRunningTasks = append(parsedRunningTasks, parsedRunningTask)

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

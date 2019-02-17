package taskmaster

import (
	"time"

	"github.com/go-ole/go-ole"
)

const (
	TASK_VALIDATE_ONLY                = 1
	TASK_CREATE                       = 2
	TASK_UPDATE                       = 4
	TASK_CREATE_OR_UPDATE             = 6
	TASK_DISABLE                      = 8
	TASK_DONT_ADD_PRINCIPAL_ACE       = 0x10
	TASK_IGNORE_REGISTRATION_TRIGGERS = 0x20
)

const (
	TASK_STATE_UNKNOWN = iota
	TASK_STATE_DISABLED
	TASK_STATE_QUEUED
	TASK_STATE_READY
	TASK_STATE_RUNNING
)

const (
	TASK_RUNLEVEL_LUA = iota
	TASK_RUNLEVEL_HIGHEST
)

const (
	TASK_ACTION_EXEC         = 0
	TASK_ACTION_COM_HANDLER  = 5
	TASK_ACTION_SEND_EMAIL   = 6
	TASK_ACTION_SHOW_MESSAGE = 7
)

const (
	TASK_LOGON_NONE = iota
	TASK_LOGON_PASSWORD
	TASK_LOGON_S4U
	TASK_LOGON_INTERACTIVE_TOKEN
	TASK_LOGON_GROUP
	TASK_LOGON_SERVICE_ACCOUNT
	TASK_LOGON_INTERACTIVE_TOKEN_OR_PASSWORD
)

const (
	TASK_COMPATIBILITY_AT = iota
	TASK_COMPATIBILITY_V1
	TASK_COMPATIBILITY_V2
	TASK_COMPATIBILITY_V2_1
	TASK_COMPATIBILITY_V2_2
	TASK_COMPATIBILITY_V2_3
	TASK_COMPATIBILITY_V2_4
)

const (
	TASK_INSTANCES_PARALLEL = iota
	TASK_INSTANCES_QUEUE
	TASK_INSTANCES_IGNORE_NEW
	TASK_INSTANCES_STOP_EXISTING
)

const (
	TASK_TRIGGER_EVENT                = 0
	TASK_TRIGGER_TIME                 = 1
	TASK_TRIGGER_DAILY                = 2
	TASK_TRIGGER_WEEKLY               = 3
	TASK_TRIGGER_MONTHLY              = 4
	TASK_TRIGGER_MONTHLYDOW           = 5
	TASK_TRIGGER_IDLE                 = 6
	TASK_TRIGGER_REGISTRATION         = 7
	TASK_TRIGGER_BOOT                 = 8
	TASK_TRIGGER_LOGON                = 9
	TASK_TRIGGER_SESSION_STATE_CHANGE = 11
	TASK_TRIGGER_CUSTOM_TRIGGER_01    = 12
)

const (
	TASK_CONSOLE_CONNECT = iota
	TASK_CONSOLE_DISCONNECT
	TASK_REMOTE_CONNECT
	TASK_REMOTE_DISCONNECT
	TASK_SESSION_LOCK
	TASK_SESSION_UNLOCK
)

type TaskService struct {
	taskServiceObj *ole.IDispatch
	isInitialized  bool
	isConnected    bool

	RootFolder      RootFolder
	RunningTasks    []*RunningTask
	RegisteredTasks []*RegisteredTask
}

type RootFolder struct {
	folderObj *ole.IDispatch
	TaskFolder
}

type TaskFolder struct {
	Name            string
	Path            string
	SubFolders      []*TaskFolder
	RegisteredTasks []*RegisteredTask
}

type RunningTask struct {
	taskObj       *ole.IDispatch
	CurrentAction string
	EnginePID     int
	InstanceGUID  string
	Name          string
	Path          string
	State         int
}

type RegisteredTask struct {
	taskObj        *ole.IDispatch
	Name           string
	Path           string
	Definition     Definition
	Enabled        bool
	State          int
	MissedRuns     int
	NextRunTime    time.Time
	LastRunTime    time.Time
	LastTaskResult int
}

type Definition struct {
	Actions          []Action
	Context          string
	Data             string
	Principal        Principal
	RegistrationInfo RegistrationInfo
	Settings         TaskSettings
	Triggers         []Trigger
	XMLText          string
}

type Action interface {
	GetID() string
	GetType() int
}

type TypeHolder struct {
	Type int
}

type TaskAction struct {
	ID string
	TypeHolder
}

type ExecAction struct {
	TaskAction
	Path       string
	Args       string
	WorkingDir string
}

type ComHandlerAction struct {
	TaskAction
	ClassID string
	Data    string
}

type EmailAction struct {
	TaskAction
	Body    string
	Server  string
	Subject string
	To      string
	Cc      string
	Bcc     string
	ReplyTo string
	From    string
}

type MessageAction struct {
	TaskAction
	Title   string
	Message string
}

type Principal struct {
	Name      string
	GroupID   string
	ID        string
	LogonType int
	RunLevel  int
	UserID    string
}

type RegistrationInfo struct {
	Author             string
	Date               string
	Description        string
	Documentation      string
	SecurityDescriptor string
	Source             string
	URI                string
	Version            string
}

type TaskSettings struct {
	AllowDemandStart       bool
	AllowHardTerminate     bool
	Compatibility          int
	DeleteExpiredTaskAfter string
	DontStartOnBatteries   bool
	Enabled                bool
	TimeLimit              string
	Hidden                 bool
	IdleSettings
	MultipleInstances int
	NetworkSettings
	Priority                 int
	RestartCount             int
	RestartInterval          string
	RunOnlyIfIdle            bool
	RunOnlyIfNetworkAvalible bool
	StartWhenAvalible        bool
	StopIfGoingOnBatteries   bool
	WakeToRun                bool
}

type IdleSettings struct {
	IdleDuration  string
	RestartOnIdle bool
	StopOnIdleEnd bool
	WaitTimeout   string
}

type NetworkSettings struct {
	ID   string
	Name string
}

type Trigger interface {
	GetEnabled() bool
	GetEndBoundary() string
	GetExecutionTimeLimit() string
	GetID() string
	GetRepitionDuration() string
	GetRepitionInterval() string
	GetStartBoundary() string
	GetStopAtDurationEnd() bool
	GetType() int
}

type TaskTrigger struct {
	Enabled            bool
	EndBoundary        string
	ExecutionTimeLimit string
	ID                 string
	RepetitionPattern
	StartBoundary string
	TypeHolder
}

// https://docs.microsoft.com/en-us/windows/desktop/api/taskschd/nn-taskschd-irepetitionpattern
type RepetitionPattern struct {
	RepitionDuration  string
	RepitionInterval  string
	StopAtDurationEnd bool
}

type BootTrigger struct {
	TaskTrigger
	Delay string
}

type DailyTrigger struct {
	TaskTrigger
	DaysInterval int
	RandomDelay  string
}

type EventTrigger struct {
	TaskTrigger
	Delay        string
	Subscription string
	ValueQueries map[string]string
}

type IdleTrigger struct {
	TaskTrigger
}

type LogonTrigger struct {
	TaskTrigger
	Delay  string
	UserID string
}

type MonthlyDOWTrigger struct {
	TaskTrigger
	DaysOfWeek           int
	MonthsOfYear         int
	RandomDelay          string
	RunOnLastWeekOfMonth bool
	WeeksOfMonth         int
}

type MonthlyTrigger struct {
	TaskTrigger
	DaysOfMonth          int
	MonthsOfYear         int
	RandomDelay          string
	RunOnLastWeekOfMonth bool
}

type RegistrationTrigger struct {
	TaskTrigger
	Delay string
}

type SessionStateChangeTrigger struct {
	TaskTrigger
	Delay       string
	StateChange int
	UserId      string
}

type TimeTrigger struct {
	TaskTrigger
	RandomDelay string
}

type WeeklyTrigger struct {
	TaskTrigger
	DaysOfWeek    int
	RandomDelay   string
	WeeksInterval int
}

type CustomTrigger struct {
	TaskTrigger
}

func (a TaskAction) GetID() string {
	return a.ID
}

func (t TypeHolder) GetType() int {
	return t.Type
}

func (t TaskTrigger) GetRepitionDuration() string {
	return t.RepitionDuration
}

func (t TaskTrigger) GetEnabled() bool {
	return t.Enabled
}

func (t TaskTrigger) GetEndBoundary() string {
	return t.EndBoundary
}

func (t TaskTrigger) GetExecutionTimeLimit() string {
	return t.ExecutionTimeLimit
}

func (t TaskTrigger) GetID() string {
	return t.ID
}

func (t TaskTrigger) GetRepitionInterval() string {
	return t.RepitionInterval
}

func (t TaskTrigger) GetStartBoundary() string {
	return t.StartBoundary
}

func (t TaskTrigger) GetStopAtDurationEnd() bool {
	return t.StopAtDurationEnd
}

func (d *Definition) AddExecAction(path, args, workingDir, id string) {
	d.Actions = append(d.Actions, ExecAction{
		Path:       path,
		Args:       args,
		WorkingDir: workingDir,
		TaskAction: TaskAction{
			ID: id,
			TypeHolder: TypeHolder{
				Type: TASK_ACTION_EXEC,
			},
		},
	})
}

func (d *Definition) AddComHandlerAction(clsid, data, id string) {
	d.Actions = append(d.Actions, ComHandlerAction{
		ClassID: clsid,
		Data:    data,
		TaskAction: TaskAction{
			ID: id,
			TypeHolder: TypeHolder{
				Type: TASK_ACTION_COM_HANDLER,
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
			TypeHolder: TypeHolder{
				Type: TASK_TRIGGER_BOOT,
			},
		},
	})
}

func (d *Definition) AddDailyTrigger(daysInterval int, randomDelay, id string, startBoundary, endBoundary time.Time, timeLimit, repitionDuration, repitionInterval string, stopAtDurationEnd, enabled bool) {
	startBoundaryStr := TimeToTaskDate(startBoundary)
	endBoundaryStr := TimeToTaskDate(endBoundary)

	d.Triggers = append(d.Triggers, DailyTrigger{
		DaysInterval: daysInterval,
		RandomDelay:  randomDelay,
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
			TypeHolder: TypeHolder{
				Type: TASK_TRIGGER_DAILY,
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
			TypeHolder: TypeHolder{
				Type: TASK_TRIGGER_EVENT,
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
			TypeHolder: TypeHolder{
				Type: TASK_TRIGGER_IDLE,
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
			TypeHolder: TypeHolder{
				Type: TASK_TRIGGER_LOGON,
			},
		},
	})
}

func (d *Definition) AddMonthlyDOWTrigger(daysOfWeek, weeksOfMonth, monthsOfYear int, runOnLastWeekOfMonth bool, randomDelay, id string, startBoundary, endBoundary time.Time, timeLimit, repitionDuration, repitionInterval string, stopAtDurationEnd, enabled bool) {
	startBoundaryStr := TimeToTaskDate(startBoundary)
	endBoundaryStr := TimeToTaskDate(endBoundary)

	d.Triggers = append(d.Triggers, MonthlyDOWTrigger{
		DaysOfWeek:           daysOfWeek,
		MonthsOfYear:         monthsOfYear,
		RandomDelay:          randomDelay,
		RunOnLastWeekOfMonth: runOnLastWeekOfMonth,
		WeeksOfMonth:         weeksOfMonth,
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
			TypeHolder: TypeHolder{
				Type: TASK_TRIGGER_MONTHLYDOW,
			},
		},
	})
}

func (d *Definition) AddMonthlyTrigger(daysOfMonth, monthsOfYear int, randomDelay, id string, startBoundary, endBoundary time.Time, timeLimit, repitionDuration, repitionInterval string, stopAtDurationEnd, enabled bool) {
	startBoundaryStr := TimeToTaskDate(startBoundary)
	endBoundaryStr := TimeToTaskDate(endBoundary)

	d.Triggers = append(d.Triggers, MonthlyTrigger{
		DaysOfMonth:  daysOfMonth,
		MonthsOfYear: monthsOfYear,
		RandomDelay:  randomDelay,
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
			TypeHolder: TypeHolder{
				Type: TASK_TRIGGER_MONTHLY,
			},
		},
	})
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
			TypeHolder: TypeHolder{
				Type: TASK_TRIGGER_REGISTRATION,
			},
		},
	})
}

func (d *Definition) AddSessionStateChangeTrigger(userID string, stateChange int, delay, id string, startBoundary, endBoundary time.Time, timeLimit, repitionDuration, repitionInterval string, stopAtDurationEnd, enabled bool) {
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
			TypeHolder: TypeHolder{
				Type: TASK_TRIGGER_SESSION_STATE_CHANGE,
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
			TypeHolder: TypeHolder{
				Type: TASK_TRIGGER_TIME,
			},
		},
	})
}

func (d *Definition) AddWeeklyTrigger(daysOfWeek, weeksInterval int, randomDelay, id string, startBoundary, endBoundary time.Time, timeLimit, repitionDuration, repitionInterval string, stopAtDurationEnd, enabled bool) {
	startBoundaryStr := TimeToTaskDate(startBoundary)
	endBoundaryStr := TimeToTaskDate(endBoundary)

	d.Triggers = append(d.Triggers, WeeklyTrigger{
		DaysOfWeek:    daysOfWeek,
		RandomDelay:   randomDelay,
		WeeksInterval: weeksInterval,
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
			TypeHolder: TypeHolder{
				Type: TASK_TRIGGER_WEEKLY,
			},
		},
	})
}

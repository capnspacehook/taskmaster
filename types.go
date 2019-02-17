package taskmaster

import (
	"time"

	"github.com/go-ole/go-ole"
)

type TaskActionType int

const (
	TASK_ACTION_EXEC         TaskActionType = 0
	TASK_ACTION_COM_HANDLER  TaskActionType = 5
	TASK_ACTION_SEND_EMAIL   TaskActionType = 6
	TASK_ACTION_SHOW_MESSAGE TaskActionType = 7
)

type TaskCompatibility int

const (
	TASK_COMPATIBILITY_AT TaskCompatibility = iota
	TASK_COMPATIBILITY_V1
	TASK_COMPATIBILITY_V2
	TASK_COMPATIBILITY_V2_1
	TASK_COMPATIBILITY_V2_2
	TASK_COMPATIBILITY_V2_3
	TASK_COMPATIBILITY_V2_4
)

type TaskCreationFlags int

const (
	TASK_VALIDATE_ONLY                TaskCreationFlags = 1
	TASK_CREATE                       TaskCreationFlags = 2
	TASK_UPDATE                       TaskCreationFlags = 4
	TASK_CREATE_OR_UPDATE             TaskCreationFlags = 6
	TASK_DISABLE                      TaskCreationFlags = 8
	TASK_DONT_ADD_PRINCIPAL_ACE       TaskCreationFlags = 0x10
	TASK_IGNORE_REGISTRATION_TRIGGERS TaskCreationFlags = 0x20
)

type TaskEnumFlags int

const (
	TASK_ENUM_HIDDEN TaskEnumFlags = 1
)

type TaskInstancesPolicy int

const (
	TASK_INSTANCES_PARALLEL TaskInstancesPolicy = iota
	TASK_INSTANCES_QUEUE
	TASK_INSTANCES_IGNORE_NEW
	TASK_INSTANCES_STOP_EXISTING
)

type TaskLogonType int

const (
	TASK_LOGON_NONE TaskLogonType = iota
	TASK_LOGON_PASSWORD
	TASK_LOGON_S4U
	TASK_LOGON_INTERACTIVE_TOKEN
	TASK_LOGON_GROUP
	TASK_LOGON_SERVICE_ACCOUNT
	TASK_LOGON_INTERACTIVE_TOKEN_OR_PASSWORD
)

type TaskRunFlags int

const (
	TASK_RUN_NO_FLAGS           TaskRunFlags = 0
	TASK_RUN_AS_SELF            TaskRunFlags = 1
	TASK_RUN_IGNORE_CONSTRAINTS TaskRunFlags = 2
	TASK_RUN_USE_SESSION_ID     TaskRunFlags = 4
	TASK_RUN_USER_SID           TaskRunFlags = 8
)

type TaskRunLevel int

const (
	TASK_RUNLEVEL_LUA TaskRunLevel = iota
	TASK_RUNLEVEL_HIGHEST
)

type TaskSessionStateChangeType int

const (
	TASK_CONSOLE_CONNECT TaskSessionStateChangeType = iota
	TASK_CONSOLE_DISCONNECT
	TASK_REMOTE_CONNECT
	TASK_REMOTE_DISCONNECT
	TASK_SESSION_LOCK
	TASK_SESSION_UNLOCK
)

type TaskState int

const (
	TASK_STATE_UNKNOWN TaskState = iota
	TASK_STATE_DISABLED
	TASK_STATE_QUEUED
	TASK_STATE_READY
	TASK_STATE_RUNNING
)

type TaskTriggerType int

const (
	TASK_TRIGGER_EVENT                TaskTriggerType = 0
	TASK_TRIGGER_TIME                 TaskTriggerType = 1
	TASK_TRIGGER_DAILY                TaskTriggerType = 2
	TASK_TRIGGER_WEEKLY               TaskTriggerType = 3
	TASK_TRIGGER_MONTHLY              TaskTriggerType = 4
	TASK_TRIGGER_MONTHLYDOW           TaskTriggerType = 5
	TASK_TRIGGER_IDLE                 TaskTriggerType = 6
	TASK_TRIGGER_REGISTRATION         TaskTriggerType = 7
	TASK_TRIGGER_BOOT                 TaskTriggerType = 8
	TASK_TRIGGER_LOGON                TaskTriggerType = 9
	TASK_TRIGGER_SESSION_STATE_CHANGE TaskTriggerType = 11
	TASK_TRIGGER_CUSTOM_TRIGGER_01    TaskTriggerType = 12
)

type TaskService struct {
	taskServiceObj        *ole.IDispatch
	isInitialized         bool
	isConnected           bool
	connectedDomain       string
	connectedComputerName string
	connectedUser         string

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
	State         TaskState
}

type RegisteredTask struct {
	taskObj        *ole.IDispatch
	Name           string
	Path           string
	Definition     Definition
	Enabled        bool
	State          TaskState
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
	GetType() TaskActionType
}

type TaskActionTypeHolder struct {
	Type TaskActionType
}

type TaskAction struct {
	ID string
	TaskActionTypeHolder
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
	LogonType TaskLogonType
	RunLevel  TaskRunLevel
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
	Compatibility          TaskCompatibility
	DeleteExpiredTaskAfter string
	DontStartOnBatteries   bool
	Enabled                bool
	TimeLimit              string
	Hidden                 bool
	IdleSettings
	MultipleInstances TaskInstancesPolicy
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
	GetType() TaskTriggerType
}

type TaskTriggerTypeHolder struct {
	Type TaskTriggerType
}

type TaskTrigger struct {
	Enabled            bool
	EndBoundary        string
	ExecutionTimeLimit string
	ID                 string
	RepetitionPattern
	StartBoundary string
	TaskTriggerTypeHolder
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
	StateChange TaskSessionStateChangeType
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

func (t TaskService) IsConnected() bool {
	return t.isConnected
}

func (t TaskService) GetConnectedDomain() string {
	return t.connectedDomain
}

func (t TaskService) GetConnectedComputerName() string {
	return t.connectedComputerName
}

func (t TaskService) GetConnectedUser() string {
	return t.connectedUser
}

func (a TaskAction) GetID() string {
	return a.ID
}

func (t TaskActionTypeHolder) GetType() TaskActionType {
	return t.Type
}

func (t TaskTriggerTypeHolder) GetType() TaskTriggerType {
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
			TaskActionTypeHolder: TaskActionTypeHolder{
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
			TaskActionTypeHolder: TaskActionTypeHolder{
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
			TaskTriggerTypeHolder: TaskTriggerTypeHolder{
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
			TaskTriggerTypeHolder: TaskTriggerTypeHolder{
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
			TaskTriggerTypeHolder: TaskTriggerTypeHolder{
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
			TaskTriggerTypeHolder: TaskTriggerTypeHolder{
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
			TaskTriggerTypeHolder: TaskTriggerTypeHolder{
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
			TaskTriggerTypeHolder: TaskTriggerTypeHolder{
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
			TaskTriggerTypeHolder: TaskTriggerTypeHolder{
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
			TaskTriggerTypeHolder: TaskTriggerTypeHolder{
				Type: TASK_TRIGGER_REGISTRATION,
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
			TaskTriggerTypeHolder: TaskTriggerTypeHolder{
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
			TaskTriggerTypeHolder: TaskTriggerTypeHolder{
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
			TaskTriggerTypeHolder: TaskTriggerTypeHolder{
				Type: TASK_TRIGGER_WEEKLY,
			},
		},
	})
}

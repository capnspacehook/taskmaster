package taskmaster

import (
	"time"

	"github.com/go-ole/go-ole"
)

const (
	TASK_VALIDATE_ONLY = iota
	TASK_CREATE
	TASK_UPDATE
	TASK_CREATE_OR_UPDATE
	TASK_DISABLE
	TASK_DONT_ADD_PRINCIPAL_ACE
	TASK_IGNORE_REGISTRATION_TRIGGERS
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
	TASK_ACTION_EXEC = 0
	TASK_ACTION_COM_HANDLER = 1
	TASK_ACTION_SEND_EMAIL = 2
	TASK_ACTION_SHOW_MESSAGE = 3
	TASK_ACTION_CUSTOM_HANDLER = 5
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
	TASK_TRIGGER_EVENT = iota
	TASK_TRIGGER_TIME
	TASK_TRIGGER_DAILY
	TASK_TRIGGER_WEEKLY
	TASK_TRIGGER_MONTHLY
	TASK_TRIGGER_MONTHLYDOW
	TASK_TRIGGER_IDLE
	TASK_TRIGGER_REGISTRATION
	TASK_TRIGGER_BOOT
	TASK_TRIGGER_LOGON
	TASK_TRIGGER_SESSION_STATE_CHANGE
	TASK_TRIGGER_CUSTOM_TRIGGER_01
)

type TaskService struct {
	taskServiceObj 	*ole.IDispatch
	isInitialized	bool
	isConnected		bool

	RunningTasks	[]RunningTask
	RegisteredTasks []RegisteredTask

}

type RunningTask struct {
	taskObj 		*ole.IDispatch
	CurrentAction	string
	EnginePID		int
	InstanceGUID	string
	Name			string
	Path			string
	State			int
}

type RegisteredTask struct {
	taskObj			*ole.IDispatch
	Name			string
	Path			string
	Definition 		Definition
	Enabled			bool
	State			int
	MissedRuns		int
	NextRunTime		time.Time
	LastRunTime		time.Time
	LastTaskResult	int
}

type Definition struct {
	actionCollectionObj		*ole.IDispatch
	triggerCollectionObj 	*ole.IDispatch
	Actions					[]Action
	Context					string
	Data					string
	Principal				Principal
	RegistrationInfo		RegistrationInfo
	Settings				TaskSettings
	Triggers				[]Trigger
	XMLText					string
}

type Action interface {
	GetType() 	int
}

type TaskAction struct {
	ID 			string
	Type		int
}

type ExecAction struct {
	TaskAction
	Path		string
	Args 		string
	WorkingDir 	string
}

type ComHandlerAction struct {
	TaskAction
	ClassID 	string
	Data		string
}

type EmailAction struct {
	TaskAction
	Body		string
	Server		string
	Subject		string
	To 			string
	Cc			string
	Bcc			string
	ReplyTo		string
	From 		string
}

type MessageAction struct {
	TaskAction
	Title 		string
	Message 	string
}

type Principal struct {
	Name		string
	GroupID		string
	ID			string
	LogonType	int
	RunLevel	int
	UserID		string
}

type RegistrationInfo struct {
	Author 				string
	Date 				string
	Description			string
	Documentation 		string
	SecurityDescriptor 	string
	Source				string
	URI 				string
	Version				string
}

type TaskSettings struct {
	AllowDemandStart			bool
	AllowHardTerminate			bool
	Compatibility				int
	DeleteExpiredTaskAfter		string
	DontStartOnBatteries		bool
	Enabled						bool
	TimeLimit					string
	Hidden						bool
	IdleSettings				IdleSettings
	MultipleInstances			int
	NetworkSettings				NetworkSettings
	Priority					int
	RestartCount				int
	RestartInterval				string
	RunOnlyIfIdle				bool
	RunOnlyIfNetworkAvalible	bool
	StartWhenAvalible			bool
	StopIfGoingOnBatteries		bool
	WakeToRun					bool
}

type IdleSettings struct {
	IdleDuration		string
	RestartOnIdle		bool
	StopOnIdleEnd		bool
	WaitTimeout			string
}

type NetworkSettings struct {
	ID 		string
	Name	string
}

type Trigger interface {
	GetType()	int
}

type TaskTrigger struct {
	Enabled				bool
	EndBoundary			string
	ExecutionTimeLimit 	string
	ID 					string
	Repetition			RepetitionPattern
	StartBoundary		string
	Type				int
}

type RepetitionPattern struct {
	Duration 			string
	Interval 			string
	StopAtDurationEnd	bool
}

type BootTrigger struct {
	TaskTrigger
	Delay 			string
}

type DailyTrigger struct {
	TaskTrigger
	DaysInterval	int
	RandomDelay		string
}

type EventTrigger struct {
	TaskTrigger
	Delay 			string
	Subscription	string
	ValueQueries	ValueQueries
}

type ValueQueries struct {
	valueQueriesObj		*ole.IDispatch
	ValueQueries		map[string]string
}

type IdleTrigger struct {
	TaskTrigger
}

type LogonTrigger struct {
	TaskTrigger
	Delay 			string
	UserID			string
}

type MonthlyDOWTrigger struct {
	TaskTrigger
	DaysOfWeek				int
	MonthsOfYear			int
	RandomDelay				string
	RunOnLastWeekOnMonth	bool
	WeeksOfMonth			int
}

type MonthlyTrigger struct {
	TaskTrigger
	DaysOfMonth				int
	MonthsOfYear			int
	RandomDelay				string
	RunOnLastWeekOnMonth	bool
}

type RegistrationTrigger struct {
	TaskTrigger
	Delay 			string
}

type TimeTrigger struct {
	TaskTrigger
	RandomDelay		string
}

type WeeklyTrigger struct {
	TaskTrigger
	DaysOfWeek		int
	RandomDelay		string
	WeeksInterval	int
}

type SessionStateChangeTrigger struct {
	TaskTrigger
}

func (e ExecAction) GetType() int {
	return e.Type
}

func (c ComHandlerAction) GetType() int {
	return c.Type
}

func (e EmailAction) GetType() int {
	return e.Type
}

func (m MessageAction) GetType() int {
	return m.Type
}

func (b BootTrigger) GetType() int {
	return b.TaskTrigger.Type
}

func (d DailyTrigger) GetType() int {
	return d.TaskTrigger.Type
}

func (e EventTrigger) GetType() int {
	return e.TaskTrigger.Type
}

func (i IdleTrigger) GetType() int {
	return i.TaskTrigger.Type
}

func (l LogonTrigger) GetType() int {
	return l.TaskTrigger.Type
}

func (m MonthlyDOWTrigger) GetType() int {
	return m.TaskTrigger.Type
}

func (m MonthlyTrigger) GetType() int {
	return m.TaskTrigger.Type
}

func (r RegistrationTrigger) GetType() int {
	return r.TaskTrigger.Type
}

func (t TimeTrigger) GetType() int {
	return t.TaskTrigger.Type
}

func (w WeeklyTrigger) GetType() int {
	return w.TaskTrigger.Type
}

func (s SessionStateChangeTrigger) GetType() int {
	return s.TaskTrigger.Type
}

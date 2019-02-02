package taskmaster

import (
	"time"

	"github.com/go-ole/go-ole"
)

const (
	TASK_STATE_UNKNOWN = iota
	TASK_STATE_DISABLED
	TASK_STATE_QUEUED
	TASK_STATE_READY
	TASK_STATE_RUNNING
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
	TASK_RUNLEVEL_LUA = iota
  	TASK_RUNLEVEL_HIGHEST
)

const (
	TASK_ACTION_EXEC = 0
	TASK_ACTION_SEND_EMAIL = 1
	TASK_ACTION_SHOW_MESSAGE =2
	TASK_ACTION_COM_HANDLER = 5
)

type TaskService struct {
	taskServiceObj 	*ole.IDispatch
	isInitialized	bool
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
	Actions				ActionCollection
	Data				string
	Principal			Principal
	RegistrationInfo	RegistrationInfo
	Settings			TaskSettings
	Triggers			[]Trigger
	XMLText				string
}

type ActionCollection struct {
	Context		string
	Actions		[]interface{}
}

type ExecAction struct {
	ID			string
	Type 		int
	Path		string
	Args 		string
	WorkingDir 	string
}

type ComHandlerAction struct {
	ID 		string
	Type 	int
	ClassID string
	Data	string
}

type EmailAction struct {
	ID		string
	Type 	int
	Body	string
	Server	string
	Subject	string
	To 		string
	Cc		string
	Bcc		string
	ReplyTo	string
	From 	string
}

type MessageAction struct {
	ID		string
	Type 	int
	Title 	string
	Message string
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

}

type TaskSettings struct {

}

type Trigger struct {

}

package taskmaster

import (
	"time"
)

const (
	TASK_STATE_UNKNOWN = iota
	TASK_STATE_DISABLED
	TASK_STATE_QUEUED
	TASK_STATE_READY
	TASK_STATE_RUNNING
)

const (
	TASK_ACTION_EXEC = 0
	TASK_ACTION_COM_HANDLER = 5
	TASK_ACTION_SEND_EMAIL
	TASK_ACTION_SHOW_MESSAGE
)

type ScheduledTask struct {
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
	Principal			Principle
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

type Principle struct {

}

type RegistrationInfo struct {

}

type TaskSettings struct {

}

type Trigger struct {

}

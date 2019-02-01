package taskmaster

import (
	"time"
)

const (
	TASK_STATE_UNKNOWN		= iota
	TASK_STATE_DISABLED		= iota
	TASK_STATE_QUEUED		= iota
	TASK_STATE_READY		= iota
	TASK_STATE_RUNNING		= iota
)

type ScheduledTask struct {
	Name			string
	Path			string
	Enabled			bool
	State			int
	MissedRuns		int
	NextRunTime		time.Time
	LastRunTime		time.Time
	LastTaskResult	int
	XML				string
}



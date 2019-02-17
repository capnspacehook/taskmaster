package taskmaster

import (
	"errors"

	//ole "github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

// Stop cancels a running task
func (r *RunningTask) Stop() error {
	stopResult := oleutil.MustCallMethod(r.taskObj, "Stop").Val
	if stopResult != 0 {
		return errors.New("cannot stop running task; access is denied")
	}

	r.taskObj.Release()

	return nil
}

func (r *RegisteredTask) Run(args []string, flags TaskRunFlags, sessionID int, user string) error {
	if sessionID != 0 {
		flags |= TASK_RUN_USE_SESSION_ID
	}

	_, err := oleutil.CallMethod(r.taskObj, "RunEx", args, int(flags), sessionID, user)
	if err != nil {
		return err
	}

	return nil
}

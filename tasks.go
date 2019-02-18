package taskmaster

import (
	"errors"

	//ole "github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

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
	if sessionID != 0 {
		flags |= TASK_RUN_USE_SESSION_ID
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

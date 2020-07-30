// +build windows

package taskmaster

import (
	"fmt"

	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

func (d *Definition) AddAction(action Action) {
	d.Actions = append(d.Actions, action)
}

func (d *Definition) AddTrigger(trigger Trigger) {
	d.Triggers = append(d.Triggers, trigger)
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

// Stop kills and releases a running task.
// https://docs.microsoft.com/en-us/windows/desktop/api/taskschd/nf-taskschd-irunningtask-stop
func (r *RunningTask) Stop() error {
	_, err := oleutil.CallMethod(r.taskObj, "Stop")
	if err != nil {
		return fmt.Errorf("error calling Stop on %s IRunningTask: %s", r.Path, err)
	}

	r.taskObj.Release()
	r.isReleased = true

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

// Run starts an instance of a registered task. If the task was started successfully,
// a pointer to a running task will be returned.
// https://docs.microsoft.com/en-us/windows/desktop/api/taskschd/nf-taskschd-iregisteredtask-run
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
// The returned slice may contain nil entries if tasks are stopped while they are being parsed.
// https://docs.microsoft.com/en-us/windows/desktop/api/taskschd/nf-taskschd-iregisteredtask-getinstances
func (r *RegisteredTask) GetInstances() ([]*RunningTask, error) {
	runningTasks, err := oleutil.CallMethod(r.taskObj, "GetInstances", 0)
	if err != nil {
		return nil, fmt.Errorf("error calling GetInstances on %s IRegisteredTask: %s", r.Path, err)
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
func (r *RegisteredTask) Stop() error {
	_, err := oleutil.CallMethod(r.taskObj, "Stop", 0)
	if err != nil {
		return err
	}

	return nil
}

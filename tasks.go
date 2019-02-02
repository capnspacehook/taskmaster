package taskmaster

import (
	"errors"
	"time"

	ole "github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

const (
	TASK_ENUM_HIDDEN = 1
)

// GetRegisteredTasks returns a list of registered scheduled tasks
func GetRegisteredTasks() ([]ScheduledTask, error) {
	var err error

	err = ole.CoInitialize(0)
	if err != nil {
		return nil, err
	}
	defer ole.CoUninitialize()

	schedClassID, err := ole.ClassIDFrom("Schedule.Service.1")
	if err != nil {
		return nil, err
	}
	taskSchedulerObj, err := ole.CreateInstance(schedClassID, nil)
	if err != nil {
		return nil, err
	}
	if taskSchedulerObj == nil {
		return nil, errors.New("Could not create ITaskService object")
	}
	defer taskSchedulerObj.Release()

	tskSchdlr := taskSchedulerObj.MustQueryInterface(ole.IID_IDispatch)
	defer tskSchdlr.Release()
	connectResult := oleutil.MustCallMethod(tskSchdlr, "Connect").Val
	if connectResult != 0 {
		switch connectResult {
		case 0x80070005:
			return nil, errors.New("access is denied to connect to the task scheduler service")
		case 0x80041315:
			return nil, errors.New("the task scheduler service is not running")
		case 0x8007000e:
			return nil, errors.New("the application does not have enough memory to complete the operation")
		case 53:
			return nil, errors.New("the computer name specified in the serverName parameter does not exist")
		case 50:
			return nil, errors.New("the remote computer does not support remote task scheduling")
		}
	}

	taskFolder := oleutil.MustCallMethod(tskSchdlr, "GetFolder", "\\").ToIDispatch()
	defer taskFolder.Release()

	taskFolderList := oleutil.MustCallMethod(taskFolder, "GetFolders", 0).ToIDispatch()
	defer taskFolderList.Release()

	var scheduledTasks []ScheduledTask

	var enumTaskFolders func(*ole.VARIANT) error
	enumTaskFolders = func (v *ole.VARIANT) error {
		taskFolder := v.ToIDispatch()
		defer taskFolder.Release()

		taskCollection := oleutil.MustCallMethod(taskFolder, "GetTasks", TASK_ENUM_HIDDEN).ToIDispatch()
		defer taskCollection.Release()

		oleutil.ForEach(taskCollection, func(v *ole.VARIANT) error {
			task := v.ToIDispatch()
			defer task.Release()

			scheduledTask, _ := parseTask(task)
			scheduledTasks = append(scheduledTasks, scheduledTask)

			return nil
		})

		taskFolderList := oleutil.MustCallMethod(taskFolder, "GetFolders", 0).ToIDispatch()
		defer taskFolderList.Release()

		oleutil.ForEach(taskFolderList, enumTaskFolders)

		return nil
	}

	err = oleutil.ForEach(taskFolderList, enumTaskFolders)
	if err != nil {
		return nil, err
	}

	return scheduledTasks, nil
}

func parseTask(task *ole.IDispatch) (ScheduledTask, error) {
	var err error

	name := oleutil.MustGetProperty(task, "Name").ToString()
	path := oleutil.MustGetProperty(task, "Path").ToString()
	enabled := oleutil.MustGetProperty(task, "Enabled").Value().(bool)
	state := int(oleutil.MustGetProperty(task, "State").Val)
	missedRuns := int(oleutil.MustGetProperty(task, "NumberOfMissedRuns").Val)
	nextRunTime := oleutil.MustGetProperty(task, "NextRunTime").Value().(time.Time)
	lastRunTime := oleutil.MustGetProperty(task, "LastRunTime").Value().(time.Time)
	lastTaskResult := int(oleutil.MustGetProperty(task, "LastTaskResult").Val)

	definition := oleutil.MustGetProperty(task, "Definition").ToIDispatch()
	defer definition.Release()
	actions := oleutil.MustGetProperty(definition, "Actions").ToIDispatch()
	defer actions.Release()
	context := oleutil.MustGetProperty(actions, "Context").ToString()

	var taskActions []interface{}
	err = oleutil.ForEach(actions, func(v *ole.VARIANT) error {
		action := v.ToIDispatch()
		defer action.Release()

		taskAction, err := parseTaskAction(action)
		if err != nil {
			return err
		}

		taskActions = append(taskActions, taskAction)

		return nil
	})
	if err != nil {
		return ScheduledTask{}, err
	}

	actionCollection := ActionCollection{
		Context:	context,
		Actions:	taskActions,
	}

	taskDef := Definition{
		Actions:	actionCollection,
	}

	scheduledTask := ScheduledTask{
		Name:			name,
		Path:			path,
		Definition:		taskDef,
		Enabled:		enabled,
		State:			state,
		MissedRuns:		missedRuns,
		NextRunTime:	nextRunTime,
		LastRunTime:	lastRunTime,
		LastTaskResult:	lastTaskResult,
	}

	return scheduledTask, nil
}

func parseTaskAction(action *ole.IDispatch) (interface{}, error) {
	id := oleutil.MustGetProperty(action, "Id").ToString()
	actionType := int(oleutil.MustGetProperty(action, "Type").Val)

	switch actionType {
	case TASK_ACTION_EXEC:
		args := oleutil.MustGetProperty(action, "Arguments").ToString()
		path := oleutil.MustGetProperty(action, "Path").ToString()
		workingDir := oleutil.MustGetProperty(action, "WorkingDirectory").ToString()

		execAction := ExecAction{
			ID:		id,
			Type:	actionType,
			Path: 	path,
			Args: 	args,
			WorkingDir: workingDir,
		}

		return execAction, nil
	case TASK_ACTION_COM_HANDLER:
		classID := oleutil.MustGetProperty(action, "ClassId").ToString()
		data := oleutil.MustGetProperty(action, "Data").ToString()

		comHandlerAction := ComHandlerAction{
			ID:			id,
			Type:		actionType,
			ClassID: 	classID,
			Data:		data,
		}

		return comHandlerAction, nil
	default:
		return nil, errors.New("unsupported IAction type")
	}

	return nil, nil
}

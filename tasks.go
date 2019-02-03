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

func (t *TaskService) initialize() error {
	var err error

	err = ole.CoInitialize(0)
	if err != nil {
		return err
	}

	schedClassID, err := ole.ClassIDFrom("Schedule.Service.1")
	if err != nil {
		return err
	}
	taskSchedulerObj, err := ole.CreateInstance(schedClassID, nil)
	if err != nil {
		return err
	}
	if taskSchedulerObj == nil {
		return errors.New("Could not create ITaskService object")
	}
	defer taskSchedulerObj.Release()

	tskSchdlr := taskSchedulerObj.MustQueryInterface(ole.IID_IDispatch)
	t.taskServiceObj = tskSchdlr
	t.isInitialized = true

	return nil
}

func (t *TaskService) Connect() error {
	if !t.isInitialized {
		err := t.initialize()
		if err != nil {
			return err
		}
	}

	connectResult := oleutil.MustCallMethod(t.taskServiceObj, "Connect").Val
	if connectResult != 0 {
		switch connectResult {
		case 0x80070005:
			return errors.New("access is denied to connect to the task scheduler service")
		case 0x80041315:
			return errors.New("the task scheduler service is not running")
		case 0x8007000e:
			return errors.New("the application does not have enough memory to complete the operation")
		case 53:
			return errors.New("the computer name specified in the serverName parameter does not exist")
		case 50:
			return errors.New("the remote computer does not support remote task scheduling")
		}
	}

	return nil
}

func (t *TaskService) Disconnect() {
	ole.CoUninitialize()
}

// GetRegisteredTasks returns a list of registered scheduled tasks
func (t *TaskService) GetRegisteredTasks() ([]RegisteredTask, error) {
	var err error

	taskFolder := oleutil.MustCallMethod(t.taskServiceObj, "GetFolder", "\\").ToIDispatch()
	defer taskFolder.Release()

	taskFolderList := oleutil.MustCallMethod(taskFolder, "GetFolders", 0).ToIDispatch()
	defer taskFolderList.Release()

	var registeredTasks []RegisteredTask

	var enumTaskFolders func(*ole.VARIANT) error
	enumTaskFolders = func (v *ole.VARIANT) error {
		taskFolder := v.ToIDispatch()
		defer taskFolder.Release()

		taskCollection := oleutil.MustCallMethod(taskFolder, "GetTasks", TASK_ENUM_HIDDEN).ToIDispatch()
		defer taskCollection.Release()

		oleutil.ForEach(taskCollection, func(v *ole.VARIANT) error {
			task := v.ToIDispatch()

			registeredTask, err := parseTask(task)
			if err != nil {
				return err
			}
			registeredTasks = append(registeredTasks, registeredTask)

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

	return registeredTasks, nil
}

func parseTask(task *ole.IDispatch) (RegisteredTask, error) {
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

	var taskActions []Action
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
		return RegisteredTask{}, err
	}

	principal := oleutil.MustGetProperty(definition, "Principal").ToIDispatch()
	defer principal.Release()
	taskPrincipal := parsePrincipal(principal)

	taskDef := Definition{
		Actions:		taskActions,
		Context:		context,
		Principal: 		taskPrincipal,
	}

	RegisteredTask := RegisteredTask{
		taskObj:		task,
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

	return RegisteredTask, nil
}

func parseTaskAction(action *ole.IDispatch) (Action, error) {
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
	case TASK_ACTION_COM_HANDLER, TASK_ACTION_CUSTOM_HANDLER:
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
}

func parsePrincipal(taskDef *ole.IDispatch) Principal {
	name := oleutil.MustGetProperty(taskDef, "DisplayName").ToString()
	groupID := oleutil.MustGetProperty(taskDef, "GroupId").ToString()
	id := oleutil.MustGetProperty(taskDef, "Id").ToString()
	logonType := int(oleutil.MustGetProperty(taskDef, "LogonType").Val)
	runLevel := int(oleutil.MustGetProperty(taskDef, "RunLevel").Val)
	userID := oleutil.MustGetProperty(taskDef, "UserId").ToString()

	principle := Principal{
		Name:		name,
		GroupID: 	groupID,
		ID:			id,
		LogonType:	logonType,
		RunLevel:	runLevel,
		UserID:		userID,
	}

	return principle
}

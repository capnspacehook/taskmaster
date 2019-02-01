package taskmaster

import (
	"errors"
	"time"

	ole "github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

const (
	TASK_ENUM_HIDDEN = 0x1
)

func handle(err error) {
	if err != nil {
		panic(err)
	}
}

func GetRegisteredTasks() ([]ScheduledTask, error) {
	var err error

	err = ole.CoInitialize(0)
	handle(err)
	defer ole.CoUninitialize()

	schedClassID, err := ole.ClassIDFrom("Schedule.Service.1")
	handle(err)
	taskSchedulerObj, err := ole.CreateInstance(schedClassID, nil)
	handle(err)
	if taskSchedulerObj == nil {
		panic("Could not create object")
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
			return nil, errors.New("The task scheduler service is not running")
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

			name := oleutil.MustGetProperty(task, "Name").ToString()
			path := oleutil.MustGetProperty(task, "Path").ToString()
			enabled := oleutil.MustGetProperty(task, "Enabled").Value().(bool)
			state := int(oleutil.MustGetProperty(task, "State").Val)
			missedRuns := int(oleutil.MustGetProperty(task, "NumberOfMissedRuns").Val)
			nextRunTime := oleutil.MustGetProperty(task, "NextRunTime").Value().(time.Time)
			lastRunTime := oleutil.MustGetProperty(task, "LastRunTime").Value().(time.Time)
			lastTaskResult := int(oleutil.MustGetProperty(task, "LastTaskResult").Val)

			scheduledTask := ScheduledTask{
				Name:			name,
				Path:			path,
				Enabled:		enabled,
				State:			state,
				MissedRuns:		missedRuns,
				NextRunTime:	nextRunTime,
				LastRunTime:	lastRunTime,
				LastTaskResult:	lastTaskResult,
			}

			scheduledTasks = append(scheduledTasks, scheduledTask)

			return nil
		})

		taskFolderList := oleutil.MustCallMethod(taskFolder, "GetFolders", 0).ToIDispatch()
		defer taskFolderList.Release()

		oleutil.ForEach(taskFolderList, enumTaskFolders)

		return nil
	}

	oleutil.ForEach(taskFolderList, enumTaskFolders)

	return scheduledTasks, nil
}

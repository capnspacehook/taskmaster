// +build windows

package taskmaster

import (
	"errors"
	"fmt"
	"os"
	"os/user"
	"strings"
	"time"

	ole "github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
	"github.com/rickb777/date/period"
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

	t.RegisteredTasks = make(map[string]*RegisteredTask)

	return nil
}

// Connect connects to a local or remote Task Scheduler service. This function
// has to be run before any other functions in taskmaster can be used. If  the
// serverName parameter is empty, a connection to the local Task Scheduler service
// will be attempted. If the user and password parameters are empty, the current
// token will be used for authentication.
func Connect(serverName, domain, username, password string) (*TaskService, error) {
	var err error
	var taskService TaskService

	if !taskService.isInitialized {
		err = taskService.initialize()
		if err != nil {
			return nil, err
		}
	}

	_, err = oleutil.CallMethod(taskService.taskServiceObj, "Connect", serverName, username, domain, password)
	if err != nil {
		errCode := GetOLEErrorCode(err)
		switch errCode {
		case 0x80070005:
			return nil, errors.New("access is denied to connect to the Task Scheduler service")
		case 0x80041315:
			return nil, errors.New("the Task Scheduler service is not running")
		case 0x8007000e:
			return nil, errors.New("the application does not have enough memory to complete the operation")
		case 0x80070032, 53:
			return nil, errors.New("cannot connect to target computer")
		case 50:
			return nil, errors.New("cannot connect to the XP or server 2003 computer")
		default:
			return nil, err
		}
	}

	if serverName == "" {
		serverName, err = os.Hostname()
		if err != nil {
			return nil, err
		}
	}
	if domain == "" {
		domain = serverName
	}
	if username == "" {
		currentUser, err := user.Current()
		if err != nil {
			return nil, err
		}
		username = strings.Split(currentUser.Username, "\\")[1]
	}
	taskService.connectedDomain = domain
	taskService.connectedComputerName = serverName
	taskService.connectedUser = username

	rootFolderObj := oleutil.MustCallMethod(taskService.taskServiceObj, "GetFolder", "\\").ToIDispatch()
	rootFolder := RootFolder{
		folderObj: rootFolderObj,
		TaskFolder: TaskFolder{
			Name: "\\",
			Path: "\\",
		},
	}
	taskService.RootFolder = rootFolder

	taskService.isConnected = true

	return &taskService, nil
}

// Cleanup frees all the Task Scheduler COM objects that have been created.
// If this function is not called before the parent program terminates,
// memory leaks will occur.
func (t *TaskService) Cleanup() {
	t.freeRunningTasks()
	t.freeRegisteredTasks()

	t.taskServiceObj.Release()
	ole.CoUninitialize()

	t.isInitialized = false
	t.isConnected = false
}

func (t *TaskService) freeRunningTasks() {
	for _, runningTask := range t.RunningTasks {
		runningTask.Release()
	}
}

func (t *TaskService) freeRegisteredTasks() {
	if t.RootFolder.folderObj != nil {
		t.RootFolder.folderObj.Release()
	}

	for _, registeredTask := range t.RegisteredTasks {
		registeredTask.taskObj.Release()
	}
}

// GetRunningTasks enumerates the Task Scheduler database for all currently running tasks.
// If run multiple times, the TaskService object will be updated to contain the current running
// tasks.
func (t *TaskService) GetRunningTasks() error {
	var err error

	// flush the stored running tasks so we can get the latest ones
	t.freeRunningTasks()
	t.RunningTasks = nil

	runningTasks := oleutil.MustCallMethod(t.taskServiceObj, "GetRunningTasks", TASK_ENUM_HIDDEN).ToIDispatch()
	defer runningTasks.Release()
	err = oleutil.ForEach(runningTasks, func(v *ole.VARIANT) error {
		task := v.ToIDispatch()

		runningTask := parseRunningTask(task)
		t.RunningTasks = append(t.RunningTasks, runningTask)

		return nil
	})
	if err != nil {
		return err
	}

	return nil
}

// GetRegisteredTasks enumerates the Task Scheduler database for all currently registered tasks.
// If run multiple times, the TaskService object will be updated to contain the current registered
// tasks.
func (t *TaskService) GetRegisteredTasks() error {
	var err error

	// if we already have registered tasks stored, flush them so we can get the lastest tasks
	if len(t.RegisteredTasks) > 0 {
		t.freeRegisteredTasks()
		t.RegisteredTasks = make(map[string]*RegisteredTask)
	}

	// get tasks from root folder
	rootTaskCollection := oleutil.MustCallMethod(t.RootFolder.folderObj, "GetTasks", int(TASK_ENUM_HIDDEN)).ToIDispatch()
	defer rootTaskCollection.Release()
	err = oleutil.ForEach(rootTaskCollection, func(v *ole.VARIANT) error {
		task := v.ToIDispatch()

		registeredTask, path, err := parseRegisteredTask(task)
		if err != nil {
			return err
		}
		t.RegisteredTasks[path] = registeredTask
		t.RootFolder.RegisteredTasks = append(t.RootFolder.RegisteredTasks, registeredTask)

		return nil
	})
	if err != nil {
		return err
	}

	taskFolderList := oleutil.MustCallMethod(t.RootFolder.folderObj, "GetFolders", 0).ToIDispatch()
	defer taskFolderList.Release()

	// recursively enumerate folders and tasks
	var initEnumTaskFolders func(*TaskFolder) func(*ole.VARIANT) error
	initEnumTaskFolders = func(parentFolder *TaskFolder) func(*ole.VARIANT) error {
		var enumTaskFolders func(*ole.VARIANT) error
		enumTaskFolders = func(v *ole.VARIANT) error {
			taskFolder := v.ToIDispatch()
			defer taskFolder.Release()

			name := oleutil.MustGetProperty(taskFolder, "Name").ToString()
			path := oleutil.MustGetProperty(taskFolder, "Path").ToString()
			taskCollection := oleutil.MustCallMethod(taskFolder, "GetTasks", int(TASK_ENUM_HIDDEN)).ToIDispatch()
			defer taskCollection.Release()

			taskSubFolder := &TaskFolder{
				Name: name,
				Path: path,
			}

			var err error
			err = oleutil.ForEach(taskCollection, func(v *ole.VARIANT) error {
				task := v.ToIDispatch()

				registeredTask, path, err := parseRegisteredTask(task)
				if err != nil {
					return err
				}
				t.RegisteredTasks[path] = registeredTask
				taskSubFolder.RegisteredTasks = append(taskSubFolder.RegisteredTasks, registeredTask)

				return nil
			})
			if err != nil {
				return err
			}

			parentFolder.SubFolders = append(parentFolder.SubFolders, taskSubFolder)

			taskFolderList := oleutil.MustCallMethod(taskFolder, "GetFolders", 0).ToIDispatch()
			defer taskFolderList.Release()

			err = oleutil.ForEach(taskFolderList, initEnumTaskFolders(taskSubFolder))
			if err != nil {
				return err
			}

			return nil
		}

		return enumTaskFolders
	}

	err = oleutil.ForEach(taskFolderList, initEnumTaskFolders(&t.RootFolder.TaskFolder))
	if err != nil {
		return err
	}

	return nil
}

// GetRegisteredTask attempts to find the specified registered task and returns a
// pointer to it if it exists. If it doesn't exist, nil will be returned in place of
// the registered task.
func (t *TaskService) GetRegisteredTask(path string) (*RegisteredTask, error) {
	taskObj, err := oleutil.CallMethod(t.RootFolder.folderObj, "GetTask", path)
	if err != nil {
		return nil, nil
	}

	task, _, err := parseRegisteredTask(taskObj.ToIDispatch())
	if err != nil {
		return nil, err
	}
	if _, exists := t.RegisteredTasks[path]; exists {
		t.RegisteredTasks[path].taskObj.Release()
	}
	t.RegisteredTasks[path] = task

	return task, nil
}

// NewTaskDefinition returns a new task definition that can be used to register a new task.
// Task settings and properties are set to Task Scheduler default values.
func (t TaskService) NewTaskDefinition() Definition {
	newDef := Definition{}

	newDef.Principal.LogonType = TASK_LOGON_INTERACTIVE_TOKEN
	newDef.Principal.RunLevel = TASK_RUNLEVEL_LUA

	newDef.RegistrationInfo.Date = time.Now()

	newDef.Settings.AllowDemandStart = true
	newDef.Settings.AllowHardTerminate = true
	newDef.Settings.Compatibility = TASK_COMPATIBILITY_V2
	newDef.Settings.DontStartOnBatteries = true
	newDef.Settings.Enabled = true
	newDef.Settings.Hidden = false
	newDef.Settings.IdleSettings.IdleDuration = period.NewHMS(0, 10, 0) // PT10M
	newDef.Settings.IdleSettings.WaitTimeout = period.NewHMS(1, 0, 0)   // PT1H
	newDef.Settings.MultipleInstances = TASK_INSTANCES_IGNORE_NEW
	newDef.Settings.Priority = 7
	newDef.Settings.RestartCount = 0
	newDef.Settings.RestartOnIdle = false
	newDef.Settings.RunOnlyIfIdle = false
	newDef.Settings.RunOnlyIfNetworkAvailable = false
	newDef.Settings.StartWhenAvailable = false
	newDef.Settings.StopIfGoingOnBatteries = true
	newDef.Settings.StopOnIdleEnd = true
	newDef.Settings.TimeLimit = period.NewHMS(72, 0, 0) // PT72H
	newDef.Settings.WakeToRun = false

	return newDef
}

func (t *TaskService) CreateTask(path string, newTaskDef Definition, overwrite bool) (*RegisteredTask, bool, error) {
	return t.CreateTaskEx(path, newTaskDef, "", "", newTaskDef.Principal.LogonType, overwrite)
}

// CreateTask creates a registered tasks on the connected computer. CreateTask returns
// true if the task was successfully registered, and false if the overwrite parameter
// is false and a task at the specified path already exists.
func (t *TaskService) CreateTaskEx(path string, newTaskDef Definition, username, password string, logonType TaskLogonType, overwrite bool) (*RegisteredTask, bool, error) {
	var err error

	if path[0] != '\\' {
		return nil, false, errors.New("path must start with root folder '\\'")
	}

	nameIndex := strings.LastIndex(path, "\\")
	folderPath := path[:nameIndex]

	if !t.taskFolderExist(folderPath) {
		oleutil.MustCallMethod(t.RootFolder.folderObj, "CreateFolder", folderPath, "")
	} else {
		if t.registeredTaskExist(path) {
			if !overwrite {
				return nil, false, nil
			}
			_, err = oleutil.CallMethod(t.RootFolder.folderObj, "DeleteTask", path, 0)
			if err != nil {
				errCode := GetOLEErrorCode(err)
				switch errCode {
				case 0x80070005:
					return nil, false, fmt.Errorf("access is denied to delete %s task", path)
				default:
					return nil, false, err
				}
			}
		}
	}

	newTaskObj, err := t.modifyTask(path, newTaskDef, username, password, logonType, TASK_CREATE)
	if err != nil {
		return nil, false, err
	}

	newTask, _, err := parseRegisteredTask(newTaskObj)
	if err != nil {
		return nil, false, err
	}

	// TODO: update taskService with possibly new folders
	t.RegisteredTasks[path] = newTask

	return newTask, true, nil
}

func (t *TaskService) UpdateTask(path string, newTaskDef Definition) (*RegisteredTask, error) {
	return t.UpdateTaskEx(path, newTaskDef, "", "", newTaskDef.Principal.LogonType)
}

// UpdateTask updates a registered task.
func (t *TaskService) UpdateTaskEx(path string, newTaskDef Definition, username, password string, logonType TaskLogonType) (*RegisteredTask, error) {
	var err error

	if path[0] != '\\' {
		return nil, errors.New("path must start with root folder '\\'")
	}

	if !t.registeredTaskExist(path) {
		return nil, errors.New("registered task doesn't exist")
	}

	newTaskObj, err := t.modifyTask(path, newTaskDef, username, password, logonType, TASK_UPDATE)
	if err != nil {
		return nil, err
	}

	// update the internal database of registered tasks
	newTask, _, err := parseRegisteredTask(newTaskObj)
	if err != nil {
		return nil, err
	}
	if updatedTask, ok := t.RegisteredTasks[path]; ok {
		updatedTask.taskObj.Release()
	}
	t.RegisteredTasks[path] = newTask

	return newTask, nil
}

func (t *TaskService) modifyTask(path string, newTaskDef Definition, username, password string, logonType TaskLogonType, flags TaskCreationFlags) (*ole.IDispatch, error) {
	var err error

	if newTaskDef.Actions == nil {
		return nil, errors.New("task must have at least one action")
	}

	if newTaskDef.Principal.UserID != "" && newTaskDef.Principal.GroupID != "" {
		return nil, errors.New("both UserId and GroupId are defined for the principal; they are mutually exclusive")
	}

	// set default UserID if UserID and GroupID both aren't set
	if newTaskDef.Principal.UserID == "" && newTaskDef.Principal.GroupID == "" {
		newTaskDef.Principal.UserID = t.connectedDomain + "\\" + t.connectedUser
	}

	newTaskDefObj := oleutil.MustCallMethod(t.taskServiceObj, "NewTask", 0).ToIDispatch()
	defer newTaskDefObj.Release()

	err = fillDefinitionObj(newTaskDef, newTaskDefObj)
	if err != nil {
		return nil, err
	}

	newTaskObj, err := oleutil.CallMethod(t.RootFolder.folderObj, "RegisterTaskDefinition", path, newTaskDefObj, int(flags), username, password, int(logonType), "")
	if err != nil {
		errCode := GetOLEErrorCode(err)
		switch errCode {
		case 0x80070005:
			return nil, errors.New("access is denied to connect to the Task Scheduler service")
		case 0x8007000e:
			return nil, errors.New("the application does not have enough memory to complete the operation")
		case 0x0004131C:
			return nil, errors.New("the task is registered, but may fail to start; batch logon privilege needs to be enabled for the task principal")
		case 0x0004131B:
			return nil, errors.New("the task is registered, but not all specified triggers will start the task")
		case 0x80041330:
			return nil, errors.New("deprecated feature used")
		default:
			return nil, err
		}
	}

	return newTaskObj.ToIDispatch(), nil
}

// DeleteFolder removes a task folder from the connected computer. If the deleteRecursively parameter
// is set to true, all tasks and subfolders will be removed recursively. If it's set to false, DeleteFolder
// will return true if the folder was empty and deleted successfully, and false otherwise.
func (t *TaskService) DeleteFolder(path string, deleteRecursively bool) (bool, error) {
	var err error

	if path[0] != '\\' {
		return false, errors.New("path must start with root folder '\\'")
	}

	if t.registeredTaskExist(path) {
		return false, errors.New("input path is a registered task, not a task folder")
	}

	taskFolder, err := oleutil.CallMethod(t.taskServiceObj, "GetFolder", path)
	if err != nil {
		return false, errors.New("task folder doesn't exist")
	}

	taskFolderObj := taskFolder.ToIDispatch()
	defer taskFolderObj.Release()
	taskCollection := oleutil.MustCallMethod(taskFolderObj, "GetTasks", int(TASK_ENUM_HIDDEN)).ToIDispatch()
	defer taskCollection.Release()
	if !deleteRecursively && oleutil.MustGetProperty(taskCollection, "Count").Val > 0 {
		return false, nil
	}

	folderCollection := oleutil.MustCallMethod(taskFolderObj, "GetFolders", int(TASK_ENUM_HIDDEN)).ToIDispatch()
	defer folderCollection.Release()
	if !deleteRecursively && oleutil.MustGetProperty(folderCollection, "Count").Val > 0 {
		return false, nil
	}

	if deleteRecursively {
		// delete tasks in parent folder
		deleteAllTasks := func(v *ole.VARIANT) error {
			taskObj := v.ToIDispatch()
			defer taskObj.Release()

			name := oleutil.MustGetProperty(taskObj, "Path").ToString()

			return t.DeleteTask(name)
		}
		err = oleutil.ForEach(taskCollection, deleteAllTasks)
		if err != nil {
			return false, err
		}

		var deleteTasksRecursively func(*ole.VARIANT) error
		deleteTasksRecursively = func(v *ole.VARIANT) error {
			var err error

			folderObj := v.ToIDispatch()
			defer folderObj.Release()

			tasks := oleutil.MustCallMethod(folderObj, "GetTasks", int(TASK_ENUM_HIDDEN)).ToIDispatch()
			defer tasks.Release()

			err = oleutil.ForEach(tasks, deleteAllTasks)
			if err != nil {
				return err
			}

			subFolders := oleutil.MustCallMethod(folderObj, "GetFolders", int(TASK_ENUM_HIDDEN)).ToIDispatch()
			defer subFolders.Release()

			err = oleutil.ForEach(subFolders, deleteTasksRecursively)
			if err != nil {
				return err
			}

			currentFolderPath := oleutil.MustGetProperty(folderObj, "Path").ToString()
			_, err = oleutil.CallMethod(t.RootFolder.folderObj, "DeleteFolder", currentFolderPath, 0)
			if err != nil {
				errCode := GetOLEErrorCode(err)
				switch errCode {
				case 0x80070005:
					return fmt.Errorf("access is denied to delete %s task", path)
				default:
					return err
				}
			}

			return nil
		}

		// delete all subfolders and tasks recursively
		err = oleutil.ForEach(folderCollection, deleteTasksRecursively)
		if err != nil {
			return false, err
		}
	}

	// delete parent folder
	_, err = oleutil.CallMethod(t.RootFolder.folderObj, "DeleteFolder", path, 0)
	if err != nil {
		return false, err
	}

	return true, nil
}

// DeleteTask removes a registered task from the connected computer.
func (t *TaskService) DeleteTask(path string) error {
	var err error

	if path[0] != '\\' {
		return errors.New("path must start with root folder '\\'")
	}

	if !t.registeredTaskExist(path) {
		return errors.New("registered task doesn't exist")
	}

	_, err = oleutil.CallMethod(t.RootFolder.folderObj, "DeleteTask", path, 0)
	if err != nil {
		errCode := GetOLEErrorCode(err)
		switch errCode {
		case 0x80070005:
			return fmt.Errorf("access is denied to delete %s task", path)
		default:
			return err
		}
	}

	// update the internal database of registered tasks
	if deletedTask, exists := t.RegisteredTasks[path]; exists {
		deletedTask.taskObj.Release()
		delete(t.RegisteredTasks, path)
	}

	return nil
}

func (t *TaskService) registeredTaskExist(path string) bool {
	_, err := oleutil.CallMethod(t.RootFolder.folderObj, "GetTask", path)
	if err != nil {
		return false
	}

	return true
}

func (t *TaskService) taskFolderExist(path string) bool {
	_, err := oleutil.CallMethod(t.taskServiceObj, "GetFolder", path)
	if err != nil {
		return false
	}

	return true
}

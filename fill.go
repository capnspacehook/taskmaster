// +build windows

package taskmaster

import (
	"errors"
	"fmt"

	ole "github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

func fillDefinitionObj(definition Definition, definitionObj *ole.IDispatch) error {
	var err error

	actionsObj := oleutil.MustGetProperty(definitionObj, "Actions").ToIDispatch()
	defer actionsObj.Release()
	oleutil.MustPutProperty(actionsObj, "Context", definition.Context)
	err = fillActionsObj(definition.Actions, actionsObj)
	if err != nil {
		return fmt.Errorf("error filling IAction objects: %s", err)
	}

	oleutil.MustPutProperty(definitionObj, "Data", definition.Data)

	principalObj := oleutil.MustGetProperty(definitionObj, "Principal").ToIDispatch()
	defer principalObj.Release()
	fillPrincipalObj(definition.Principal, principalObj)

	regInfoObj := oleutil.MustGetProperty(definitionObj, "RegistrationInfo").ToIDispatch()
	defer regInfoObj.Release()
	fillRegistrationInfoObj(definition.RegistrationInfo, regInfoObj)

	settingsObj := oleutil.MustGetProperty(definitionObj, "Settings").ToIDispatch()
	defer settingsObj.Release()
	fillTaskSettingsObj(definition.Settings, settingsObj)

	triggersObj := oleutil.MustGetProperty(definitionObj, "Triggers").ToIDispatch()
	defer triggersObj.Release()
	err = fillTaskTriggersObj(definition.Triggers, triggersObj)
	if err != nil {
		return fmt.Errorf("error filling ITrigger objects: %s", err)
	}

	return nil
}
func fillActionsObj(actions []Action, actionsObj *ole.IDispatch) error {
	for _, action := range actions {
		actionType := action.GetType()
		if !checkActionType(actionType) {
			return errors.New("invalid action type")
		}

		actionObj := oleutil.MustCallMethod(actionsObj, "Create", int(actionType)).ToIDispatch()
		actionObj.Release()
		oleutil.MustPutProperty(actionObj, "Id", action.GetID())

		switch actionType {
		case TASK_ACTION_EXEC:
			execAction := action.(ExecAction)
			exeActionObj := actionObj.MustQueryInterface(ole.NewGUID("{4c3d624d-fd6b-49a3-b9b7-09cb3cd3f047}"))
			defer exeActionObj.Release()

			oleutil.MustPutProperty(exeActionObj, "Arguments", execAction.Args)
			oleutil.MustPutProperty(exeActionObj, "Path", execAction.Path)
			oleutil.MustPutProperty(exeActionObj, "WorkingDirectory", execAction.WorkingDir)
		case TASK_ACTION_COM_HANDLER:
			comHandlerAction := action.(ComHandlerAction)
			comHandlerActionObj := actionObj.MustQueryInterface(ole.NewGUID("{6d2fd252-75c5-4f66-90ba-2a7d8cc3039f}"))
			defer comHandlerActionObj.Release()

			oleutil.MustPutProperty(comHandlerActionObj, "ClassId", comHandlerAction.ClassID)
			oleutil.MustPutProperty(comHandlerActionObj, "Data", comHandlerAction.Data)
		}
	}

	return nil
}

func checkActionType(actionType TaskActionType) bool {
	switch actionType {
	case TASK_ACTION_EXEC:
		fallthrough
	case TASK_ACTION_COM_HANDLER:
		return true
	default:
		return false
	}
}

func fillPrincipalObj(principal Principal, principalObj *ole.IDispatch) {
	oleutil.MustPutProperty(principalObj, "DisplayName", principal.Name)
	oleutil.MustPutProperty(principalObj, "GroupId", principal.GroupID)
	oleutil.MustPutProperty(principalObj, "Id", principal.ID)
	oleutil.MustPutProperty(principalObj, "LogonType", int(principal.LogonType))
	oleutil.MustPutProperty(principalObj, "RunLevel", int(principal.RunLevel))
	oleutil.MustPutProperty(principalObj, "UserId", principal.UserID)
}

func fillRegistrationInfoObj(regInfo RegistrationInfo, regInfoObj *ole.IDispatch) {
	oleutil.MustPutProperty(regInfoObj, "Author", regInfo.Author)
	oleutil.MustPutProperty(regInfoObj, "Date", TimeToTaskDate(regInfo.Date))
	oleutil.MustPutProperty(regInfoObj, "Description", regInfo.Description)
	oleutil.MustPutProperty(regInfoObj, "Documentation", regInfo.Documentation)
	oleutil.MustPutProperty(regInfoObj, "SecurityDescriptor", regInfo.SecurityDescriptor)
	oleutil.MustPutProperty(regInfoObj, "Source", regInfo.Source)
	oleutil.MustPutProperty(regInfoObj, "URI", regInfo.URI)
	oleutil.MustPutProperty(regInfoObj, "Version", regInfo.Version)
}

func fillTaskSettingsObj(settings TaskSettings, settingsObj *ole.IDispatch) {
	oleutil.MustPutProperty(settingsObj, "AllowDemandStart", settings.AllowDemandStart)
	oleutil.MustPutProperty(settingsObj, "AllowHardTerminate", settings.AllowHardTerminate)
	oleutil.MustPutProperty(settingsObj, "Compatibility", int(settings.Compatibility))
	oleutil.MustPutProperty(settingsObj, "DeleteExpiredTaskAfter", settings.DeleteExpiredTaskAfter)
	oleutil.MustPutProperty(settingsObj, "DisallowStartIfOnBatteries", settings.DontStartOnBatteries)
	oleutil.MustPutProperty(settingsObj, "Enabled", settings.Enabled)
	oleutil.MustPutProperty(settingsObj, "ExecutionTimeLimit", PeriodToString(settings.TimeLimit))
	oleutil.MustPutProperty(settingsObj, "Hidden", settings.Hidden)

	idlesettingsObj := oleutil.MustGetProperty(settingsObj, "IdleSettings").ToIDispatch()
	defer idlesettingsObj.Release()
	oleutil.MustPutProperty(idlesettingsObj, "IdleDuration", PeriodToString(settings.IdleSettings.IdleDuration))
	oleutil.MustPutProperty(idlesettingsObj, "RestartOnIdle", settings.IdleSettings.RestartOnIdle)
	oleutil.MustPutProperty(idlesettingsObj, "StopOnIdleEnd", settings.IdleSettings.StopOnIdleEnd)
	oleutil.MustPutProperty(idlesettingsObj, "WaitTimeout", PeriodToString(settings.IdleSettings.WaitTimeout))

	oleutil.MustPutProperty(settingsObj, "MultipleInstances", int(settings.MultipleInstances))

	networksettingsObj := oleutil.MustGetProperty(settingsObj, "NetworkSettings").ToIDispatch()
	defer networksettingsObj.Release()
	oleutil.MustPutProperty(networksettingsObj, "Id", settings.NetworkSettings.ID)
	oleutil.MustPutProperty(networksettingsObj, "Name", settings.NetworkSettings.Name)

	oleutil.MustPutProperty(settingsObj, "Priority", settings.Priority)
	oleutil.MustPutProperty(settingsObj, "RestartCount", settings.RestartCount)
	oleutil.MustPutProperty(settingsObj, "RestartInterval", PeriodToString(settings.RestartInterval))
	oleutil.MustPutProperty(settingsObj, "RunOnlyIfIdle", settings.RunOnlyIfIdle)
	oleutil.MustPutProperty(settingsObj, "RunOnlyIfNetworkAvailable", settings.RunOnlyIfNetworkAvailable)
	oleutil.MustPutProperty(settingsObj, "StartWhenAvailable", settings.StartWhenAvailable)
	oleutil.MustPutProperty(settingsObj, "StopIfGoingOnBatteries", settings.StopIfGoingOnBatteries)
	oleutil.MustPutProperty(settingsObj, "WakeToRun", settings.WakeToRun)
}

func fillTaskTriggersObj(triggers []Trigger, triggersObj *ole.IDispatch) error {
	for _, trigger := range triggers {
		triggerType := trigger.GetType()
		if !checkTriggerType(triggerType) {
			return errors.New("invalid trigger type")
		}
		triggerObj := oleutil.MustCallMethod(triggersObj, "Create", int(triggerType)).ToIDispatch()
		defer triggerObj.Release()

		oleutil.MustPutProperty(triggerObj, "Enabled", trigger.GetEnabled())
		oleutil.MustPutProperty(triggerObj, "EndBoundary", TimeToTaskDate(trigger.GetEndBoundary()))
		oleutil.MustPutProperty(triggerObj, "ExecutionTimeLimit", PeriodToString(trigger.GetExecutionTimeLimit()))
		oleutil.MustPutProperty(triggerObj, "Id", trigger.GetID())

		repetitionObj := oleutil.MustGetProperty(triggerObj, "Repetition").ToIDispatch()
		defer repetitionObj.Release()
		oleutil.MustPutProperty(repetitionObj, "Duration", PeriodToString(trigger.GetRepetitionDuration()))
		oleutil.MustPutProperty(repetitionObj, "Interval", PeriodToString(trigger.GetRepetitionInterval()))
		oleutil.MustPutProperty(repetitionObj, "StopAtDurationEnd", trigger.GetStopAtDurationEnd())

		oleutil.MustPutProperty(triggerObj, "StartBoundary", TimeToTaskDate(trigger.GetStartBoundary()))

		switch triggerType {
		case TASK_TRIGGER_BOOT:
			bootTrigger := trigger.(BootTrigger)
			bootTriggerObj := triggerObj.MustQueryInterface(ole.NewGUID("{2a9c35da-d357-41f4-bbc1-207ac1b1f3cb}"))
			defer bootTriggerObj.Release()

			oleutil.MustPutProperty(bootTriggerObj, "Delay", PeriodToString(bootTrigger.Delay))
		case TASK_TRIGGER_DAILY:
			dailyTrigger := trigger.(DailyTrigger)
			dailyTriggerObj := triggerObj.MustQueryInterface(ole.NewGUID("{126c5cd8-b288-41d5-8dbf-e491446adc5c}"))
			defer dailyTriggerObj.Release()

			oleutil.MustPutProperty(dailyTriggerObj, "DaysInterval", int(dailyTrigger.DayInterval))
			oleutil.MustPutProperty(dailyTriggerObj, "RandomDelay", PeriodToString(dailyTrigger.RandomDelay))
		case TASK_TRIGGER_EVENT:
			eventTrigger := trigger.(EventTrigger)
			eventTriggerObj := triggerObj.MustQueryInterface(ole.NewGUID("{d45b0167-9653-4eef-b94f-0732ca7af251}"))
			defer eventTriggerObj.Release()

			oleutil.MustPutProperty(eventTriggerObj, "Delay", PeriodToString(eventTrigger.Delay))
			oleutil.MustPutProperty(eventTriggerObj, "Subscription", eventTrigger.Subscription)
			valueQueriesObj := oleutil.MustGetProperty(eventTriggerObj, "ValueQueries").ToIDispatch()
			defer valueQueriesObj.Release()

			for name, value := range eventTrigger.ValueQueries {
				oleutil.MustCallMethod(valueQueriesObj, "Create", name, value)
			}
		case TASK_TRIGGER_IDLE:
			idleTriggerObj := triggerObj.MustQueryInterface(ole.NewGUID("{d537d2b0-9fb3-4d34-9739-1ff5ce7b1ef3}"))
			defer idleTriggerObj.Release()
		case TASK_TRIGGER_LOGON:
			logonTrigger := trigger.(LogonTrigger)
			logonTriggerObj := triggerObj.MustQueryInterface(ole.NewGUID("{72dade38-fae4-4b3e-baf4-5d009af02b1c}"))
			defer logonTriggerObj.Release()

			oleutil.MustPutProperty(logonTriggerObj, "Delay", PeriodToString(logonTrigger.Delay))
			oleutil.MustPutProperty(logonTriggerObj, "UserId", logonTrigger.UserID)
		case TASK_TRIGGER_MONTHLYDOW:
			monthlyDOWTrigger := trigger.(MonthlyDOWTrigger)
			monthlyDOWTriggerObj := triggerObj.MustQueryInterface(ole.NewGUID("{77d025a3-90fa-43aa-b52e-cda5499b946a}"))
			defer monthlyDOWTriggerObj.Release()

			oleutil.MustPutProperty(monthlyDOWTriggerObj, "DaysOfWeek", int(monthlyDOWTrigger.DaysOfWeek))
			oleutil.MustPutProperty(monthlyDOWTriggerObj, "MonthsOfYear", int(monthlyDOWTrigger.MonthsOfYear))
			oleutil.MustPutProperty(monthlyDOWTriggerObj, "RandomDelay", PeriodToString(monthlyDOWTrigger.RandomDelay))
			oleutil.MustPutProperty(monthlyDOWTriggerObj, "RunOnLastWeekOfMonth", monthlyDOWTrigger.RunOnLastWeekOfMonth)
			oleutil.MustPutProperty(monthlyDOWTriggerObj, "WeeksOfMonth", int(monthlyDOWTrigger.WeeksOfMonth))
		case TASK_TRIGGER_MONTHLY:
			monthlyTrigger := trigger.(MonthlyTrigger)
			monthlyTriggerObj := triggerObj.MustQueryInterface(ole.NewGUID("{97c45ef1-6b02-4a1a-9c0e-1ebfba1500ac}"))
			defer monthlyTriggerObj.Release()

			oleutil.MustPutProperty(monthlyTriggerObj, "DaysOfMonth", int(monthlyTrigger.DaysOfMonth))
			oleutil.MustPutProperty(monthlyTriggerObj, "MonthsOfYear", int(monthlyTrigger.MonthsOfYear))
			oleutil.MustPutProperty(monthlyTriggerObj, "RandomDelay", PeriodToString(monthlyTrigger.RandomDelay))
			oleutil.MustPutProperty(monthlyTriggerObj, "RunOnLastDayOfMonth", monthlyTrigger.RunOnLastWeekOfMonth)
		case TASK_TRIGGER_REGISTRATION:
			registrationTrigger := trigger.(RegistrationTrigger)
			registrationTriggerObj := triggerObj.MustQueryInterface(ole.NewGUID("{4c8fec3a-c218-4e0c-b23d-629024db91a2}"))
			defer registrationTriggerObj.Release()

			oleutil.MustPutProperty(registrationTriggerObj, "Delay", PeriodToString(registrationTrigger.Delay))
		case TASK_TRIGGER_TIME:
			timeTrigger := trigger.(TimeTrigger)
			timeTriggerObj := triggerObj.MustQueryInterface(ole.NewGUID("{b45747e0-eba7-4276-9f29-85c5bb300006}"))
			defer timeTriggerObj.Release()

			oleutil.MustPutProperty(timeTriggerObj, "RandomDelay", PeriodToString(timeTrigger.RandomDelay))
		case TASK_TRIGGER_WEEKLY:
			weeklyTrigger := trigger.(WeeklyTrigger)
			weeklyTriggerObj := triggerObj.MustQueryInterface(ole.NewGUID("{5038fc98-82ff-436d-8728-a512a57c9dc1}"))
			defer weeklyTriggerObj.Release()

			oleutil.MustPutProperty(weeklyTriggerObj, "DaysOfWeek", int(weeklyTrigger.DaysOfWeek))
			oleutil.MustPutProperty(weeklyTriggerObj, "RandomDelay", PeriodToString(weeklyTrigger.RandomDelay))
			oleutil.MustPutProperty(weeklyTriggerObj, "WeeksInterval", int(weeklyTrigger.WeekInterval))
		case TASK_TRIGGER_SESSION_STATE_CHANGE:
			sessionStateChangeTrigger := trigger.(SessionStateChangeTrigger)
			sessionStateChangeTriggerObj := triggerObj.MustQueryInterface(ole.NewGUID("{754da71b-4385-4475-9dd9-598294fa3641}"))
			defer sessionStateChangeTriggerObj.Release()

			oleutil.MustPutProperty(sessionStateChangeTriggerObj, "Delay", PeriodToString(sessionStateChangeTrigger.Delay))
			oleutil.MustPutProperty(sessionStateChangeTriggerObj, "StateChange", int(sessionStateChangeTrigger.StateChange))
			oleutil.MustPutProperty(sessionStateChangeTriggerObj, "UserId", sessionStateChangeTrigger.UserId)
			// need to find GUID
			/*case TASK_TRIGGER_CUSTOM_TRIGGER_01:
			return nil*/
		}
	}

	return nil
}

func checkTriggerType(triggerType TaskTriggerType) bool {
	switch triggerType {
	case TASK_TRIGGER_BOOT:
		fallthrough
	case TASK_TRIGGER_DAILY:
		fallthrough
	case TASK_TRIGGER_EVENT:
		fallthrough
	case TASK_TRIGGER_IDLE:
		fallthrough
	case TASK_TRIGGER_LOGON:
		fallthrough
	case TASK_TRIGGER_MONTHLYDOW:
		fallthrough
	case TASK_TRIGGER_MONTHLY:
		fallthrough
	case TASK_TRIGGER_REGISTRATION:
		fallthrough
	case TASK_TRIGGER_TIME:
		fallthrough
	case TASK_TRIGGER_WEEKLY:
		fallthrough
	case TASK_TRIGGER_SESSION_STATE_CHANGE:
		return true
		// need to find GUID
		/*case TASK_TRIGGER_CUSTOM_TRIGGER_01:
		return nil*/
	default:
		return false
	}
}

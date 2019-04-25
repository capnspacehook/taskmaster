// +build windows

package taskmaster

import (
	"errors"
	"math"
	"time"

	"github.com/go-ole/go-ole"
	"github.com/rickb777/date/period"
)

var taskDateFormat = "2006-01-02T15:04:05"
var taskDateFormatWTimeZone = "2006-01-02T15:04:05-07:00"

func IntToDayOfMonth(dayOfMonth int) (DayOfMonth, error) {
	if dayOfMonth < 1 || dayOfMonth > 32 {
		return 0, errors.New("invalid day of month")
	}

	return DayOfMonth(math.Exp2(float64(dayOfMonth - 1))), nil
}

func TimeToTaskDate(t time.Time) string {
	defaultTime := time.Time{}
	if t == defaultTime {
		return ""
	}

	return t.Format(taskDateFormat)
}

func TaskDateToTime(s string) (time.Time, error) {
	if s == "" {
		return time.Time{}, nil
	}

	// try parsing with the first format, and it that doesn't work, use the second one.
	// this is necessary because Microsoft feels the need to randomly pick one of two
	// time formats when creating built-in tasks.
	t, err := time.Parse(taskDateFormat, s)
	if err != nil {
		t, err = time.Parse(taskDateFormatWTimeZone, s)
		if err != nil {
			return time.Time{}, err
		}
	}

	return t, nil
}

func StringToPeriod(s string) (period.Period, error) {
	if s == "" {
		return period.Period{}, nil
	}

	return period.Parse(s)
}

func PeriodToString(p period.Period) string {
	s := p.String()
	if s == "P0D" {
		return ""
	}

	return s
}

func GetOLEErrorCode(err error) uint32 {
	return err.(*ole.OleError).SubError().(ole.EXCEPINFO).SCODE()
}

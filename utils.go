package taskmaster

import (
	//"strconv"
	//"strings"
	"errors"
	"math"
	"time"
)

var taskDateFormat = "2006-01-02T15:04:05.0000000"

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
	t, err := time.Parse(taskDateFormat, s)
	if err != nil {
		return time.Time{}, err
	}

	return t, nil
}

/*func TimeToDuration(t time.Duration) string {

}

func DurationToTime(s string) time.Duration {
	index := 0
	var duration time.Duration
	var years, months, days, hours, minutes, seconds int

	if strings.Contains(s, "T") {
		if yearIndex := strings.Index(s, "Y"); yearIndex != -1 {
			years, _ = strconv.Atoi(s[1:yearIndex])
		}
		if monthIndex := strings.Index(s, "M"); monthIndex != -1 {
			months, _ = strconv.Atoi(s[1:monthIndex])
		}
		if dayIndex := strings.Index(s, "D"); dayIndex != -1 {
			days, _ = strconv.Atoi(s[1:dayIndex])
		}
	}
	if hourIndex := strings.Index(s, "H"); hourIndex != -1 {
		hours, _ = strconv.Atoi(s[1:hourIndex])
	}
	if minuteIndex := strings.Index(s, "M"); minuteIndex != -1 {
		hours, _ = strconv.Atoi(s[1:minuteIndex])
	}
	if secondIndex := strings.Index(s, "S"); secondIndex != -1 {
		seconds, _ = strconv.Atoi(s[1:secondIndex])
	}

	duration += int64(years) * time.Minute

}*/

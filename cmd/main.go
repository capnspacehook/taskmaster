package main

import (
	"fmt"

	"github.com/capnspacehook/taskmaster"
)

func main() {
	var err error

	tasks, err := taskmaster.GetRegisteredTasks()
	if err != nil {
		panic(err)
	}

	for _, task := range(tasks) {
		fmt.Printf("Name: %s\n", task.Name)
		fmt.Printf("Path: %s\n", task.Path)
		fmt.Printf("Enabled: %t\n", task.Enabled)
		fmt.Printf("Number of Missed Runs: %d\n", task.MissedRuns)
		fmt.Printf("Next Run Time: %s\n", task.NextRunTime)
		fmt.Printf("Last Run Time: %s\n", task.LastRunTime)
	}
}

package main

import (
	"fmt"
	"strings"

	"github.com/capnspacehook/taskmaster"
)

func printTasks(taskFolder taskmaster.TaskFolder, depth int) {
	padding := strings.Repeat("\t", depth)
	fmt.Println(padding + taskFolder.Path)
	for _, task := range(taskFolder.RegisteredTasks) {
		fmt.Println(padding + "\t" + task.Name)
	}

	for _, folder := range(taskFolder.SubFolders) {
		printTasks(*folder, depth + 1)
	}
}

func main() {
	var err error
	var taskService taskmaster.TaskService
	err = taskService.Connect()
	if err != nil {
		panic(err)
	}
	defer taskService.Cleanup()

	err = taskService.GetRegisteredTasks()
	if err != nil {
		panic(err)
	}

	printTasks(taskService.RootFolder, 0)
}

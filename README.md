[![Go Report Card](https://goreportcard.com/badge/github.com/capnspacehook/taskmaster)](https://goreportcard.com/report/github.com/capnspacehook/taskmaster)
[![GoDoc](https://godoc.org/github.com/capnspacehook/taskmaster?status.svg)](https://godoc.org/github.com/capnspacehook/taskmaster)

# taskmaster
Windows Task Scheduler Library for Go

![taskmaster villian](img/taskmaster.jpg "Taskmaster")

# What is taskmaster?

Taskmaster is a library for managing Scheduled Tasks in Windows. It allows you to easily create, modify, delete, execute, kill, and view scheduled tasks, on your local machine or on a remote one. It provides much more speed and power than using the native Task Scheduler GUI in Windows, and the Scheduled Task Powershell cmdlets. 

Because taskmaster interfaces directly with Task Scheduler COM objects, it allows you to do things you can't do with the Task Scheduler GUI or Powershell cmdlets. COM handler task actions can be viewed, manipulated, and created, more settings can be used when creating or modifying scheduled tasks, ect. Taskmaster exposes the full potential of Windows Scheduled Tasks in a clean, simple interface. 

# Install

To expiriment with taskmaster, compile and run the example programs.

``` shell
go get github.com/capnspacehook/taskmaster

cd /path/to/taskmaster/examples/createTimedTask
go run timedTask.go
```

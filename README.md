# Serial connection for Windows Terminal

## Overview

Microsoft's Windows Terminal is excellent terminal software.
However, there is a problem of not being able to make serial 
connections easily. So I developed a script program to make 
a serial connection in PowerShell.

This script can connect to the serial console of Linux. 
This is created by using PowerShell, and it is worked 
on Windows Terminal. The supported protocol are 
COM port and Named pipe.


## Sample movie

[![COM port connection for Windows Terminal](https://img.youtube.com/vi/ssHAWzYcqng/0.jpg)](https://www.youtube.com/watch?v=ssHAWzYcqng "COM port connection")


[![Named pipe connection for Windows Terminal](https://img.youtube.com/vi/8Pyr64KRw-c/0.jpg)](https://www.youtube.com/watch?v=8Pyr64KRw-c "COM port connection")




## Supported Tool

Microsoft Windows 10 and 11

PowerShell 7

*NOTE*

Not supported Windows PowerShell 5.1

## Usage

### SYNOPSIS

Serial connection for Windows Terminal

### DESCRIPTION

This script can connect to the serial console of Linux. 
This is created by using PowerShell, and it is worked 
on Windows Terminal. The supported protocol are 
COM port and Named pipe.

### PARAMETER port

Serial port name or Named pipe name.

### PARAMETER baudrate

The serial baud rate.

### PARAMETER pipe

Named pipe is enabled.
Specify this option when you use the named pipe.

### PARAMETER abortkey

How to close your session in the following:

"ctrlc"     CTRL+C
"ctrlb"     CTRL+b
"tildedot"  ~.

### PARAMETER conntimeout

Named pipe only.

The maximum time to wait for a successful connection 
in milliseconds.

### PARAMETER connretry

Trying to connection repeats infinitely even if 
the session is aborted.

### PARAMETER encode

Host encoding and decoding.

### PARAMETER logstart

Start logging.

### PARAMETER logverbose

Debug message will be shown.


### EXAMPLE

PS> .\serial_terminal.ps1 -port com1

PS> .\serial_terminal.ps1 -port com1 -baudrate 9600  -abortkey ctrlc

PS> .\serial_terminal.ps1 -pipe -port com1 -conntimeout 1000


## History

*09/25/2022 Ver1.0*
- First release


## License

MIT

[Please find attached file](LICENSE)

## Author

(C) 2022 Yutaka Hirata(YOULAB)

GitHub: https://github.com/yutakakn

Homepage: https://hp.vector.co.jp/authors/VA013320/index.html


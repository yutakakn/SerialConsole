# SPDX-License-Identifier: MIT
#
# Serial connection for Windows Terminal
#
# This script can connect to the serial console of Linux. 
# This is created by using PowerShell, and it is 
# worked on Windows Terminal.
# The supported protocol are COM port and Named pipe.
# 
# (C) 2022 Yutaka Hirata(YOULAB)
# https://github.com/yutakakn
# https://hp.vector.co.jp/authors/VA013320/index.html
#

<#
.SYNOPSIS

Serial connection for Windows Terminal

.DESCRIPTION

This script can connect to the serial console of Linux. 
This is created by using PowerShell, and it is worked 
on Windows Terminal. The supported protocol are 
COM port and Named pipe.

.PARAMETER port

Serial port name or Named pipe name.

.PARAMETER baudrate

The serial baud rate.

.PARAMETER pipe

Named pipe is enabled.
Specify this option when you use the named pipe.

.PARAMETER abortkey

How to close your session in the following:

"ctrlc"     CTRL+C
"ctrlb"     CTRL+b
"tildedot"  ~.

.PARAMETER conntimeout

Named pipe only.

The maximum time to wait for a successful connection 
in milliseconds.

.PARAMETER connretry

Trying to connection repeats infinitely even if 
the session is aborted.

.PARAMETER encode

Host encoding and decoding.

.PARAMETER logstart

Start logging.

.PARAMETER logverbose

Debug message will be shown.

.INPUTS
.OUTPUTS

.EXAMPLE

PS> .\serial_terminal.ps1 -port com1

.EXAMPLE

PS> .\serial_terminal.ps1 -port com1 -baudrate 9600  -abortkey ctrlc

.EXAMPLE

PS> .\serial_terminal.ps1 -pipe -port com1 -conntimeout 1000

#>

Param(
    [Parameter(Position=0, Mandatory=$true)] 
    [string] 
    $port = "COM1",

    [Parameter(Mandatory=$false)] 
    [int] 
    $baudrate = 38400,

    [Parameter(Mandatory=$false)] 
    [switch] 
    $pipe = $false,

    [Parameter(Mandatory=$false)] 
    [string] 
    $abortkey = "tildedot",

    [Parameter(Mandatory=$false)] 
    [int] 
    $conntimeout = 0,

    [Parameter(Mandatory=$false)] 
    [switch] 
    $connretry = $false,

    [Parameter(Mandatory=$false)] 
    [string] 
    $encode = "UTF-8",

    [Parameter(Mandatory=$false)] 
    [switch] 
    $logstart = $false,

    [Parameter(Mandatory=$false)] 
    [switch] 
    $logverbose = $false
)

# Strict grammar checks
Set-StrictMode -Version Latest

<#
Abort key definition for your session.

[NOTICE]
The same definition applies to the script block 
which is implemented the thread.
#>
enum AbortKeyTypes {
    ctrlc
    ctrlb
    tildedot    
}


#
# Command line parameters class
#
class CommandLineParameters
{
    [string]static $PortName 
    [int]static $BaudRate 
    [bool]static $NamedPipeFlag 
    [AbortKeyTypes]static $AbortKey 
    [int]static $ConnectTimeout
    [bool]static $ConnectRetry
    [string]static $HostEncoding
    [bool]static $LogStart
    [bool]static $Verbose

    [System.Text.Encoding]GetEncodeType([string]$encodeString) {
        [System.Text.Encoding]$retEncode = $null

        switch ($encodeString) {
            ("UTF-8") {
                $retEncode = [System.Text.Encoding]::UTF8
            }
            ("UTF8") {
                $retEncode = [System.Text.Encoding]::UTF8
            }
            ("ASCII") {
                $retEncode = [System.Text.Encoding]::ASCII
            }
            ("Default") {
                $retEncode = [System.Text.Encoding]::Default
            }
            ("SJIS") {
                $retEncode = [System.Text.Encoding]::Default
            }
            ("Shift_JIS") {
                $retEncode = [System.Text.Encoding]::Default
            }
            default {
                $retEncode = [System.Text.Encoding]::Default
            }
        }

        return $retEncode
    }

    [string]GetAborkeyDesc() {
        $str = $null
        switch ([CommandLineParameters]::AbortKey) {
            ([AbortKeyTypes]::ctrlc) {
                $str = "CTRL+C"
                break
            }
            ([AbortKeyTypes]::ctrlb) {
                $str = "CTRL+B"
                break
            }
            ([AbortKeyTypes]::tildedot) {
                $str = "~."
                break
            }
            default {
                $str = ""
            }
        }
        return $str
    }

    ShowOption() {
        Write-Host "Port: " ([CommandLineParameters]::PortName)
        Write-Host "BaudRate: " ([CommandLineParameters]::BaudRate)
        Write-Host "Namedpipe: " ([CommandLineParameters]::NamedPipeFlag)
        Write-Host "AbortKey: " ([CommandLineParameters]::AbortKey)
        Write-Host "ConnectTimeout: " ([CommandLineParameters]::ConnectTimeout)
        Write-Host "ConnectRetry: " ([CommandLineParameters]::ConnectRetry)
        Write-Host "HostEncoding: " ([CommandLineParameters]::HostEncoding)
        Write-Host "LogStart: " ([CommandLineParameters]::LogStart)
        Write-Host "Verbose: " ([CommandLineParameters]::Verbose)
    }
}

#
# Emulated abstract class
#
# [NOTICE]
# You can not directly create the instance of this class.
#
class AbstractSerialConnection 
{
    [bool]$retry
    [int]$readWaitTime    # milli seconds
    [bool]$KeyInterruptAction
    [bool]$Verbose;
    
    AbstractSerialConnection() {
        $type = $this.GetType()

        if ($type -eq [AbstractSerialConnection]) {
            throw("Class $type can not be created own instance.")
        }

        $this.retry = $false
        $this.readWaitTime = 200   
        $this.KeyInterruptAction = $false
    }

    [bool]Connect() {
        throw("Connect must be overrided.")
    }

    [bool]IsConnected() {
        throw("IsConnected must be overrided.")
    }

    [string]ReadAsync() {
        throw("ReadAsync must be overrided.")
    }

    WriteKeyData([char]$key) {
        throw("WriteKeyData must be overrided.")
    }

    DisConnect() {
        throw("DisConnect must be overrided.")
    }

    SetKeyInterruptAction() {
        $this.KeyInterruptAction = $true
    }

    [bool]GetKeyInterruptAction() {
        return $this.KeyInterruptAction
    }

    <#
    Thread function by using Script Block.
    
    [NOTTICE]
    This thread will slow down when the thread calls a method
    of the thread generator via "this" pointer. 
    Basically, I decided not to user "this" pointer except
    in special cases.
    #>
    $scriptBlockFunc = {
        # Below definition is the same as the enum for global scope.
        enum AbortKeyTypes {
            ctrlc
            ctrlb
            tildedot    
        }
        
        [bool]$firstCharFound
        $keyInfo

        $thisPtr = $args[0]
        $thisAbortKey = $args[1]
 
        for (;;) {
            if ([Console]::KeyAvailable) {
                $key = [system.console]::readkey($true)
            } else {
                <#
                if (-Not $thisPtr.IsConnected()) {
                    if ([CommandLineParameters]::Verbose) {
                        Write-Host "readKey Disconnected "
                    }
                    break
                }
                #>

                Start-Sleep -Milliseconds 1
                continue
            }

            $keyInfo = $key
            #write-host "modi " $keyInfo.modifiers "key " $keyInfo.key

            $ret = $false
            switch ($thisAbortKey) {
                ([AbortKeyTypes]::ctrlc) {
                    if (($keyInfo.modifiers -band [consolemodifiers]"control") -and 
                        ($keyInfo.key -eq "C")) {
                            $ret = $true
                    }
                }
    
                ([AbortKeyTypes]::ctrlb) {
                    if (($keyInfo.modifiers -band [consolemodifiers]"control") -and 
                        ($keyInfo.key -eq "B")) {
                            $ret = $true
                    }
                }
    
                ([AbortKeyTypes]::tildedot) {
                    # Check always if the character is a tilde(~).
                    if (
                        ($keyInfo.modifiers -band [consolemodifiers]"shift") -and 
                         (($keyInfo.key -eq [ConsoleKey]::Oem3) -or 
                          ($keyInfo.key -eq [ConsoleKey]::Oem7))
                        ) {
                        $firstCharFound = $true
                        if ([CommandLineParameters]::Verbose) {
                            Write-Host "Found ~ character."
                        }
                    } else {
                        if ($firstCharFound) {
                            $firstCharFound = $false
                            # Check if second character is a period.
                            if ($keyInfo.key -eq [ConsoleKey]::OemPeriod) {
                                $ret = $true
                                if ([CommandLineParameters]::Verbose) {
                                    Write-Host "Found . character."
                                }
                            }
                        }
                    }
                }
    
                default {
                    $ret = $false
                }
            }
    
            if ($ret) {
                Write-Host "Terminating this session by using your key abortion." -ForegroundColor Red
                $thisPtr.SetKeyInterruptAction()
                break
            }

            switch ($key.Key) {
                ([ConsoleKey]::UpArrow) {$ch = [char]0x10}    # Ctrl+P
                ([ConsoleKey]::DownArrow) {$ch = [char]0x0e}  # Ctrl+N
                ([ConsoleKey]::LeftArrow) {$ch = [char]0x02}  # Ctrl+B
                ([ConsoleKey]::RightArrow) {$ch = [char]0x06}  # Ctrl+F
                default {$ch = $key.KeyChar}
            }            

            #write-host "$ch" | format-hex
            $thisPtr.WriteKeyData($ch)
        }
    }   

    #
    # Main routine
    #
    DoMain($MyHost) {
        $ret = $this.Connect()
        if (-Not $ret) {
            return
        }

        $keystr = [CommandLineParameters]::new().GetAborkeyDesc()
        $msg = "How to close your session:`n" +
               "Type '" + $keystr + "' on the terminal."
        Write-Host $msg -ForegroundColor DarkGreen

        [console]::TreatControlCAsInput = $true

        $arglist = @($this, [CommandLineParameters]::AbortKey)
        $job = Start-ThreadJob -ArgumentList $arglist -ScriptBlock $this.scriptBlockFunc -StreamingHost $MyHost

        :READLOOP for (;;) {
            $str = $null

            for (;;) {
                if (-Not $this.IsConnected()) {
                    if ([CommandLineParameters]::Verbose) {
                        Write-Host "Your connection has been lost." $this.Handle
                    }
                    break READLOOP
                }

                $str = $this.ReadAsync()

                if ([string]::IsNullOrEmpty($str)) {
                    Start-Sleep -Milliseconds $this.readWaitTime 
                } else {
                    break
                }

                if ($this.KeyInterruptAction) {
                    if ([CommandLineParameters]::Verbose) {
                        write-host "Terminating thread..."
                    }
                    break READLOOP
                }
            }

            Write-Host -NoNewline $str
        }

        #Write-Host "job " $job.State
        Stop-job $job
        #Write-Host "job2 " $job.State

        $null = wait-job $job
        remove-job $job

        [console]::TreatControlCAsInput = $false

        $this.DisConnect()
    }

    #
    # Entry point
    #
    Main($MyHost) {
        # Start logging
        if ([CommandLineParameters]::LogStart) {
            Start-Transcript
        }

        do {
            $this.DoMain($MyHost)
            if ($this.KeyInterruptAction) {
                break
            }
            Start-Sleep -Milliseconds 1000
        } while ([CommandLineParameters]::ConnectRetry) 

        # Stop logging
        if ([CommandLineParameters]::LogStart) {
            Stop-Transcript
        }
    }

}

#
# COM port connection class
#
class ComPortConnection : AbstractSerialConnection
{
    $Handle

    ComPortConnection() {
        $this.Handle = $null
    }

    [bool]Connect() {
        $comport = [CommandLineParameters]::PortName
        $baudrate = [CommandLineParameters]::BaudRate

        try {
            $myport = New-Object System.IO.Ports.SerialPort `
            $comport, $baudrate, ([System.IO.Ports.Parity]::None)

            $myport.Encoding = [System.Text.Encoding]::GetEncoding([CommandLineParameters]::HostEncoding)

            $myport.Open()
        } catch {
            $msgOn = $false
            if ([CommandLineParameters]::Verbose) {
                $msgOn = $true
            } else {
                if ([CommandLineParameters]::ConnectRetry) {
                    $msgOn = $false
                } else {
                    $msgOn = $true
                }
            }

            if ($msgOn) {
                Write-Host "Can not open $comport for some reason." -ForegroundColor red
                Write-Host "Available ports are in the following:"
                $s = [System.IO.Ports.SerialPort]::getportnames()
                Write-Host $s
            }
            return $false
        } 

        $this.Handle = $myport
        Write-Host "Com port($comport) was opened."
        return $true
    }

    [bool]IsConnected() {
        $ret = $false
        if ($this.Handle) {
            $ret = $this.Handle.IsOpen
        }
        return $ret
    }

    [string]ReadAsync() {
        $str = $this.Handle.ReadExisting()
        return $str
    }

    WriteKeyData([char]$key) {
        if ($this.Handle) {
            $this.Handle.Write($key)
        }
    }

    DisConnect() {
        if ($this.Handle) {
            $this.Handle.Dispose()
            $this.Handle = $null

            Write-Host "Com port was disconnected."
        }
    }    
}


#
# Named pipe connection class
#
class NamedPipeConnection : AbstractSerialConnection
{
    $Handle
    [byte[]]$recvBuf
    [int]$recvBufMax
    [byte[]]$inBuffer
    [int]$inBufferMax
    [int]$lastinBufferIndex
    [System.Text.Encoding]$EncodeType
    [String]$EncodeString

    NamedPipeConnection() {
        $this.Handle = $null
        $this.recvBufMax = 1024
        #$this.recvBufMax = 16
        $this.recvBuf = New-Object byte[] $this.recvBufMax
        $this.inBufferMax = 4096
        #$this.inBufferMax = 64
        $this.inBuffer = New-Object byte[] $this.inBufferMax
        $this.lastinBufferIndex = 0
        $this.EncodeString = [CommandLineParameters]::HostEncoding
        $this.EncodeType = [CommandLineParameters]::new().GetEncodeType($this.EncodeString)
    }

    awaitTask($myTask) {
        while (-not $myTask.AsyncWaitHandle.WaitOne(200)) { 
            # none
        }
        $null = $myTask.GetAwaiter().GetResult()
    }

    [bool]Connect() {
        $PipeName = [CommandLineParameters]::PortName
        $ConnectTimeout = [CommandLineParameters]::ConnectTimeout

        try {
            $PipeHandle = New-Object -TypeName System.IO.Pipes.NamedPipeClientStream `
                -ArgumentList ".", $PipeName, 
                ([System.IO.Pipes.PipeDirection]::InOut), 
                ([System.IO.Pipes.PipeOptions]::Asynchronous)

            if ($ConnectTimeout -gt 0) {
                $task = $PipeHandle.ConnectAsync($ConnectTimeout)   
            } else {
                $task = $PipeHandle.ConnectAsync()   
            }
            $this.awaitTask($task)

        } catch {
            $msgOn = $false
            if ([CommandLineParameters]::Verbose) {
                $msgOn = $true
            } else {
                if ([CommandLineParameters]::ConnectRetry) {
                    $msgOn = $false
                } else {
                    $msgOn = $true
                }
            }

            if ($msgOn) {
                Write-Host "Can not open \\.\pipe\$PipeName for some reason." -ForegroundColor red
            }
            return $false
        } 

        $this.Handle = $PipeHandle
        Write-Host "Named pipe(\\.\pipe\$PipeName) was opened."
        return $true
    }

    [bool]IsConnected() {
        $ret = $false
        if ($this.Handle) {
            $ret = $this.Handle.IsConnected
        }
        return $ret
    }

    InsertGetBuffer($buf, $len, [ref]$retString) {
        $copied = $false
        $encStr = $null

        #$buf[0] = 0xe3
        #$len = 1

        # Check if buffer over flow
        if ($this.lastinBufferIndex + $len -ge $this.inBufferMax) {
            $tmpBufLen = $this.lastinBufferIndex + $len
            $tmpBuf = [byte[]]::new($tmpBufLen)
            [System.Buffer]::BlockCopy($this.inBuffer, 0, $tmpBuf, 0, $this.lastinBufferIndex)
            [System.Buffer]::BlockCopy($buf, 0, $tmpBuf, $this.lastinBufferIndex, $len)

            $encStr = $this.EncodeType.GetString($tmpBuf, 0, $tmpBufLen)      

            #Write-Host "Buffer over flow!" -ForegroundColor red
            #Write-Host $encStr

        } else {
            [System.Buffer]::BlockCopy($buf, 0, $this.inBuffer, $this.lastinBufferIndex, $len)
            $this.lastinBufferIndex += $len
            $copied = $true
        }

        if ($copied) {
            $decoderException = [Text.Encoding]::GetEncoding( 
                $this.EncodeString,
                (New-Object Text.EncoderReplacementFallback),
                (New-Object Text.DecoderExceptionFallback) 
            )        
    
            #
            # Step1: I will try to see if every data can be decoded.
            #
            try {
                $encStr = $decoderException.GetString($this.inBuffer, 0, $this.lastinBufferIndex)              
            } catch {
                $encStr = $null
            }

            if ($null -ne $encStr) {
                $this.lastinBufferIndex = 0

            } else {
                $lastIndex = $this.lastinBufferIndex
                $newBufLen = $orgBufLen = $lastIndex
                $localDecoder = $this.EncodeType.GetDecoder()
                $numChars = $localDecoder.GetCharCount($this.inBuffer, 0, $this.lastinBufferIndex)
                <#
                Write-Host "Start---"
                Write-Host "numChars " $numChars
                Write-Host "lastIndex " $lastIndex
                $hexString = ($this.inBuffer|ForEach-Object ToString X2) -join ' '
                Write-Host $hexString
                #>

                while ($lastIndex -ge 2) {  # >= 2
                    $localDecoder.Reset()
                    $lastIndex--
                    $num = $localDecoder.GetCharCount($this.inBuffer, 0, $lastIndex)
                    #Write-Host "num dec" $num
                    if ($num -ne $numChars) {
                        $newBufLen = $lastIndex + 1
                        break
                    }
                }
                <#
                Write-Host "Start2"
                Write-Host "lastIndex " $lastIndex
                Write-Host "newBufLen " $newBufLen
                #>

                if ($newBufLen -le $orgBufLen) { # <=
                    $encStr = $this.EncodeType.GetString($this.inBuffer, 0, $newBufLen)      
                    #Write-Host $encStr

                    $remainLen = $orgBufLen - $newBufLen
                    #Write-Host "rem " $remainLen
                    #Write-Host "lastIndex " $this.lastinBufferIndex
                    if ($remainLen -gt 0) { # > 0
                        [System.Buffer]::BlockCopy($this.inBuffer, $newBufLen, $this.inBuffer, 0, $remainLen)
                        $this.lastinBufferIndex = $remainLen
                    } else {
                        $this.lastinBufferIndex = 0
                    }

                } else {
                    $encStr = $null
                    #Write-Host "null " $lastIndex -ForegroundColor red
                    #Write-Host "newBufLen " $newBufLen -ForegroundColor red
                    #Write-Host "orgBufLen " $orgBufLen -ForegroundColor red
                }

                <#
                Write-Host "End"
                Write-Host "lastIndex " $this.lastinBufferIndex
                $hexString = ($this.inBuffer|ForEach-Object ToString X2) -join ' '
                Write-Host $hexString
                #>
            }
        }

        $retString.Value = $encStr
    }

    [string]ReadAsync() {
        $retStr = $null
        $pipe = $this.Handle
        $buf = $this.recvBuf
        $buflen = $this.recvBuf.Length
        $task = $pipe.ReadAsync($buf, 0, $buflen)  
        $intr = $false
        while (-Not $task.AsyncWaitHandle.WaitOne(200)) {
            $intr = ([AbstractSerialConnection]$this).GetKeyInterruptAction()
            if ($intr) {
                break
            }
        }

        if ($intr) {
            $retStr = $null
        } else {
            $len = $task.GetAwaiter().GetResult()
            $this.InsertGetBuffer($buf, $len, [ref]$retStr)
        }

        return $retStr
    }

    WriteKeyData([char]$key) {
        $keysArray = [char[]]::new(1)
        $keysArray[0] = $key
    
        $keysBytes = $this.EncodeType.GetBytes($keysArray)
        #write-host $keysBytes | format-hex
    
        #$this.Handle.Write($keysBytes, 0, $keysBytes.Length)
        $task = $this.Handle.WriteAsync($keysBytes, 0, $keysBytes.Length)
        $this.awaitTask($task)
    }

    DisConnect() {
        if ($this.Handle) {
            $this.Handle.Dispose()
            $this.Handle = $null

            Write-Host "Named pipe was disconnected."
        }
    }    
}


#
# Parse command line options.
#
Function CheckCommandLineParameters() {
    [CommandLineParameters]::PortName = $port
    [CommandLineParameters]::BaudRate = $baudrate
    [CommandLineParameters]::NamedPipeFlag = $pipe

    switch ($abortkey) {
        "ctrlc" { $type = [AbortKeyTypes]::ctrlc }
        "ctrlb" { $type = [AbortKeyTypes]::ctrlb }
        "tildedot" { $type = [AbortKeyTypes]::tildedot }
        default { $type = [AbortKeyTypes]::tildedot }
    }
    [CommandLineParameters]::AbortKey = $type

    [CommandLineParameters]::ConnectTimeout = $conntimeout
    [CommandLineParameters]::ConnectRetry = $connretry
    [CommandLineParameters]::HostEncoding = $encode
    [CommandLineParameters]::LogStart = $logstart
    [CommandLineParameters]::Verbose = $logverbose

    if ([CommandLineParameters]::Verbose) {
        [CommandLineParameters]::new().ShowOption()
    }
}

#
# First entry
#
CheckCommandLineParameters

if ([CommandLineParameters]::NamedPipeFlag) {
    [AbstractSerialConnection[]] $entry = @([NamedPipeConnection]::New())
} else {
    [AbstractSerialConnection[]] $entry = @([ComPortConnection]::New())
}

ForEach($obj in $entry) {
    $obj.Main($Host)
}

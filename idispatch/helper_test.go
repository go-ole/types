package idispatch

import (
	"fmt"
	"log"
	"os"
	"strings"
	"time"
	"unsafe"

	syscall "golang.org/x/sys/windows"

	"github.com/gonuts/commander"

	"github.com/go-ole/com"
	"github.com/go-ole/iconnectionpointcontainer"
	"github.com/go-ole/iunknown"
)

func Example_excel() {
	const delay = 2000000000
	com.CoInitialize()
	defer com.CoUninitialize()

	var unknown *iunknown.Unknown
	var excel *Dispatch

	com.CreateObject("Excel.Application", &unknown)
	unknown.QueryInterface(com.IDispatchInterfaceID, &excel)

	PutProperty(excel, "Visible", true)
	workbooks := MustGetProperty(excel, "Workbooks").ToIDispatch()
	workbook := MustCallMethod(workbooks, "Add", nil).ToIDispatch()
	worksheet := MustGetProperty(workbook, "Worksheets", 1).ToIDispatch()
	cell := MustGetProperty(worksheet, "Cells", 1, 1).ToIDispatch()
	PutProperty(cell, "Value", 12345)

	time.Sleep(delay)

	PutProperty(workbook, "Saved", true)
	CallMethod(workbook, "Close", false)
	CallMethod(excel, "Quit")
	excel.Release()
}

func Example_ie() {
	const delay = 2000000000
	com.CoInitialize()
	defer com.CoUninitialize()

	var unknown *iunknown.Unknown
	var ie *Dispatch

	com.CreateObject("InternetExplorer.Application", &unknown)
	unknown.QueryInterface(com.IDispatchInterfaceID, &ie)

	CallMethod(ie, "Navigate", "http://www.google.com")
	PutProperty(ie, "Visible", true)
	for {
		if MustGetProperty(ie, "Busy").Val == 0 {
			break
		}
	}

	time.Sleep(delay)

	document := MustGetProperty(ie, "document").ToIDispatch()
	window := MustGetProperty(document, "parentWindow").ToIDispatch()
	// set 'golang' to text box.
	MustCallMethod(window, "eval", "document.getElementsByName('q')[0].value = 'golang'")
	// click btnG.
	btnG := MustCallMethod(window, "eval", "document.getElementsByName('btnG')[0]").ToIDispatch()
	MustCallMethod(btnG, "click")
}

func Example_mediaplayer() {
	com.CoInitialize()
	defer com.CoUninitialize()

	var unknown *iunknown.Unknown
	var wmp *Dispatch

	err := com.CreateObject("WMPlayer.OCX", &unknown)
	if err != nil {
		log.Fatal(err)
	}

	unknown.MustQueryInterface(com.IDispatchInterfaceID, &wmp)

	collection := MustGetProperty(wmp, "MediaCollection").ToIDispatch()
	list := MustCallMethod(collection, "getAll").ToIDispatch()
	count := int(MustGetProperty(list, "count").Val)
	for i := 0; i < count; i++ {
		item := MustGetProperty(list, "item", i).ToIDispatch()
		name := MustGetProperty(item, "name").ToString()
		sourceURL := MustGetProperty(item, "sourceURL").ToString()
		fmt.Println(name, sourceURL)
	}
}

func Example_msagent() {
	com.CoInitialize()
	defer com.CoUninitialize()

	var unknown *iunknown.Unknown
	var agent *Dispatch

	err := com.CreateObject("Agent.Control.1", &unknown)
	unknown.QueryInterface(com.IDispatchInterfaceID, &agent)

	PutProperty(agent, "Connected", true)
	characters := MustGetProperty(agent, "Characters").ToIDispatch()
	CallMethod(characters, "Load", "Merlin", "c:\\windows\\msagent\\chars\\Merlin.acs")
	character := MustCallMethod(characters, "Character", "Merlin").ToIDispatch()
	CallMethod(character, "Show")
	CallMethod(character, "Speak", "こんにちわ世界")
}

func Example_msxml_rssreader() {
	com.CoInitialize()
	defer com.CoUninitialize()

	var unknown *iunknown.Unknown
	var xmlhttp *Dispatch

	err := com.CreateObject("Microsoft.XMLHTTP", &unknown)
	unknown.QueryInterface(com.IDispatchInterfaceID, &xmlhttp)
	defer xmlhttp.Release()

	MustCallMethod(xmlhttp, "open", "GET", "http://rss.slashdot.org/Slashdot/slashdot", false)
	MustCallMethod(xmlhttp, "send", nil)

	state := -1
	for state != 4 {
		state = int(MustGetProperty(xmlhttp, "readyState").Val)
		time.Sleep(10000000)
	}

	responseXml := MustGetProperty(xmlhttp, "responseXml").ToIDispatch()
	items := MustCallMethod(responseXml, "selectNodes", "/rss/channel/item").ToIDispatch()
	defer items.Release()
	length := int(MustGetProperty(items, "length").Val)

	for n := 0; n < length; n++ {
		item := MustGetProperty(items, "item", n).ToIDispatch()
		title := MustCallMethod(item, "selectSingleNode", "title").ToIDispatch()
		link := MustCallMethod(item, "selectSingleNode", "link").ToIDispatch()

		fmt.Println(MustGetProperty(title, "text").ToString())
		fmt.Println("  " + MustGetProperty(link, "text").ToString())

		title.Release()
		link.Release()
		item.Release()
	}
}

func Example_outlook() {
	com.CoInitialize()
	defer com.CoUninitialize()

	var unknown *iunknown.Unknown
	var outlook *Dispatch

	err := com.CreateObject("Outlook.Application", &unknown)
	unknown.QueryInterface(com.IDispatchInterfaceID, &outlook)
	defer outlook.Release()

	ns := MustCallMethod(outlook, "GetNamespace", "MAPI").ToIDispatch()
	folder := MustCallMethod(ns, "GetDefaultFolder", 10).ToIDispatch()
	contacts := MustCallMethod(folder, "Items").ToIDispatch()
	count := MustGetProperty(contacts, "Count").Value().(int32)

	for i := 1; i <= int(count); i++ {
		item, err := GetProperty(contacts, "Item", i)
		if err == nil && item.VariantType == com.IDispatchVariantType {
			if value, err := GetProperty(item.ToIDispatch(), "FullName"); err == nil {
				fmt.Println(value.Value())
			}
		}
	}

	MustCallMethod(outlook, "Quit")
}

func Example_itunes() {
	com.CoInitialize()
	defer com.CoUninitialize()

	var err error
	var unknown *iunknown.Unknown
	var itunes *Dispatch

	err = com.CreateObject("iTunes.Application", &unknown)
	if err != nil {
		log.Fatal(err)
	}
	err = unknown.QueryInterface(com.IDispatchInterfaceID, &itunes)
	if err != nil {
		log.Fatal(err)
	}
	defer itunes.Release()

	command := &commander.Command{
		UsageLine: os.Args[0],
		Short:     "itunes cmd",
	}

	command.Subcommands = []*commander.Command{}
	for _, name := range []string{"Play", "Stop", "Pause", "Quit"} {
		command.Subcommands = append(command.Subcommands, &commander.Command{
			Run: func(cmd *commander.Command, args []string) error {
				_, err := CallMethod(itunes, name)
				return err
			},
			UsageLine: strings.ToLower(name),
		})
	}

	err = command.Dispatch(os.Args[1:])
	if err != nil {
		log.Fatal(err)
	}
}

func Example_winsock() {
	queryInterface := func(self *interface{}, interfaceID *com.GUID, client **interface{}) uint32 {
		code := iunknown.QueryInterfaceService(self, interfaceID, *client)
		if code == com.NoInterfaceErrorCode {
			s := com.StringFromClassID(interfaceID)
			if s == "{248DD893-BB45-11CF-9ABC-0080C7E7B78D}" {
				iunknown.AddRefService(self)
				*client = self
				return com.SuccessResponseCode
			}
		}
		return com.NoInterfaceErrorCode
	}

	getIDsOfNames := func(this *iunknown.IUnknown, iid *com.GUID, wnames []*uint16, namelen int, lcid int, pdisp []int32) uintptr {
		for n := 0; n < namelen; n++ {
			pdisp[n] = int32(n)
		}
		return uintptr(com.SuccessResponseCode)
	}

	invoke := func(this *IDispatch, dispid int32, riid *com.GUID, lcid int, flags int16, dispparams *com.DISPPARAMS, result *com.Variant, pexcepinfo *com.EXCEPINFO, nerr *uint) {
		switch dispid {
		case 0:
			log.Println("DataArrival")
			winsock := (*com.EventReceiver)(unsafe.Pointer(this)).host
			var data com.Variant
			com.VariantInit(&data)
			CallMethod(winsock, "GetData", &data)
			s := string(data.ToArray().ToByteArray())
			println()
			println(s)
			println()
		case 1:
			log.Println("Connected")
			winsock := (*com.EventReceiver)(unsafe.Pointer(this)).host
			oleutil.CallMethod(winsock, "SendData", "GET / HTTP/1.0\r\n\r\n")
		case 3:
			log.Println("SendProgress")
		case 4:
			log.Println("SendComplete")
		case 5:
			log.Println("Close")
			this.Release()
		case 6:
			log.Fatal("Error")
		default:
			log.Println(dispid)
		}
		return com.NotImplementedErrorCode
	}

	com.CoInitialize()
	defer com.CoUninitialize()

	var unknown *iunknown.Unknown
	var winsock *Dispatch

	err := com.CreateObject("{248DD896-BB45-11CF-9ABC-0080C7E7B78D}", &unknown)
	if err != nil {
		panic(err.Error())
	}
	unknown.QueryInterface(com.IDispatchInterfaceID, &winsock)
	defer winsock.Release()

	classID, _ := com.ClassIDFromString("{248DD893-BB45-11CF-9ABC-0080C7E7B78D}")

	destination := &com.EventReceiver{}
	destination.VirtualTable = &VirtualTable{
		QueryInterface:   syscall.NewCallback(queryInterface),
		AddRef:           syscall.NewCallback(iunknown.AddRefService),
		Release:          syscall.NewCallback(iunknown.ReleaseService),
		GetTypeInfoCount: syscall.NewCallback(GetTypeInfoCountService),
		GetTypeInfo:      syscall.NewCallback(GetTypeInfoService),
		GetIDsOfNames:    syscall.NewCallback(getIDsOfNames),
		Invoke:           syscall.NewCallback(invoke)}
	destination.Host = winsock

	iconnectionpointcontainer.ConnectObject(winsock, classID, (*iunknown.Unknown)(unsafe.Pointer(destination)))
	_, err = CallMethod(winsock, "Connect", "127.0.0.1", 80)
	if err != nil {
		log.Fatal(err)
	}

	var m com.Msg
	for dest.ReferenceCount != 0 {
		com.GetMessage(&m, 0, 0, 0)
		com.DispatchMessage(&m)
	}
}

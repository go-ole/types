package iconnectionpoint

import (
	"unsafe"

	"github.com/go-ole/com"
	"github.com/go-ole/iunknown"
)

// InterfaceID is IID for IConnectionPoint
var InterfaceID = &com.GUID{0xB196B286, 0xBAB4, 0x101A, [8]byte{0xB6, 0x9C, 0x00, 0xAA, 0x00, 0x34, 0x1D, 0x07}}

// ContainerInterfaceID is IID for IConnectionPointContainer
var ContainerInterfaceID = &com.GUID{0xB196B284, 0xBAB4, 0x101A, [8]byte{0xB6, 0x9C, 0x00, 0xAA, 0x00, 0x34, 0x1D, 0x07}}

// ConnectionPointContainer is the structure for storing raw virtual table.
type ConnectionPointContainer iunknown.Unknown

// ConnectionPoint is the structure for storing raw virtual table.
type ConnectionPoint iunknown.Unknown

// VirtualTable is the Virtual Table for IConnectionPoint COM objects.
//
// Stores pointers to IConnectionPoint methods to be called using Syscall or
// unsafe.
type VirtualTable struct {
	iunknown.VirtualTable
	GetConnectionInterface      uintptr
	GetConnectionPointContainer uintptr
	Advise                      uintptr
	Unadvise                    uintptr
	EnumConnections             uintptr
}

// ContainerVirtualTable is the Virtual Table for IConnectionPointContainer COM
// objects.
//
// Stores pointers to IConnectionPointContainer methods to be called using
// Syscall or unsafe.
type ContainerVirtualTable struct {
	iunknown.VirtualTable
	EnumConnectionPoints uintptr
	FindConnectionPoint  uintptr
}

type IConnectionPoint interface {
	GetConnectionInterface() (*com.GUID, error)
	GetConnectionPointContainer() (*ConnectionPointContainer, error)
	Advise(unknown *iunknown.IUnknown) (uint32, error)
	Unadvise(cookie uint32) error
	EnumConnections() (element *interface{}, err error)
}

type IConnectionPointContainer interface {
	EnumConnectionPoints() (interface{}, error)
	FindConnectionPoint(interfaceID *com.GUID) (*ConnectionPoint, error)
}

// VTable converts ConnectionPoint to IConnectionPoint VirtualTable.
func (c *ConnectionPoint) VTable() *VirtualTable {
	return (*VirtualTable)(unsafe.Pointer(c.RawVTable))
}

// QueryInterface will query interface and return object by reference.
func (c *ConnectionPoint) QueryInterface(interfaceID *com.GUID, element interface{}) error {
	return iunknown.QueryInterface(c, c.VTable().QueryInterface, interfaceID, &element)
}

// AddRef increments the reference counter for object.
func (c *ConnectionPoint) AddRef() int32 {
	return iunknown.AddRef(c, c.VTable().AddRef)
}

// Release decrements the reference counter for object.
func (c *ConnectionPoint) Release() int32 {
	return iunknown.Release(c, c.VTable().Release)
}

func (c *ConnectionPoint) GetConnectionInterface() (*com.GUID, error) {
	return GetConnectionInterface(c, c.VTable().GetConnectionInterface)
}

func (c *ConnectionPoint) GetConnectionPointContainer() (*ConnectionPointContainer, error) {
	return GetConnectionPointContainer(c, c.VTable().GetConnectionPointContainer)
}

func (c *ConnectionPoint) Advise(unknown *iunknown.IUnknown) (uint32, error) {
	return Advise(c, c.VTable().Advise, unknown)
}

func (c *ConnectionPoint) Unadvise(cookie uint32) error {
	return Unadvise(c, c.VTable().Unadvise, cookie)
}

func (c *ConnectionPoint) EnumConnections() (element *interface{}, err error) {
	return EnumConnections(c, c.VTable().EnumConnections)
}

// VTable converts ConnectionPointContainer to IConnectionPointContainer
// VirtualTable.
func (c *ConnectionPointContainer) VTable() *ContainerVirtualTable {
	return (*ContainerVirtualTable)(unsafe.Pointer(c.RawVTable))
}

// QueryInterface will query interface and return object by reference.
func (c *ConnectionPointContainer) QueryInterface(interfaceID *com.GUID, element interface{}) error {
	return iunknown.QueryInterface(c, c.VTable().QueryInterface, interfaceID, &element)
}

// AddRef increments the reference counter for object.
func (c *ConnectionPointContainer) AddRef() int32 {
	return iunknown.AddRef(c, c.VTable().AddRef)
}

// Release decrements the reference counter for object.
func (c *ConnectionPointContainer) Release() int32 {
	return iunknown.Release(c, c.VTable().Release)
}

func (c *ConnectionPointContainer) EnumConnectionPoints() (interface{}, error) {
	return EnumConnectionPoints(c, c.VTable().EnumConnectionPoints)
}

func (c *ConnectionPointContainer) FindConnectionPoint(interfaceID *com.GUID) (*ConnectionPoint, error) {
	return FindConnectionPoint(c, c.VTable().FindConnectionPoint, interfaceID)
}

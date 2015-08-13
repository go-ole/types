# COM IUnknown Interface

**Unstable. Experimental status.**

[![GoDoc](https://godoc.org/github.com/go-ole/iunknown?status.svg)](https://godoc.org/github.com/go-ole/iunknown)

Cvery COM interface inherits from IUknown interface and therefore has the
following methods:

 * AddRef()
 * QueryInterface()
 * Release()

These methods are the first entries on all of the virtual tables for COM
objects.

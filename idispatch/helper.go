package idispatch

import (
	"github.com/go-ole/com"
)

// MethodAt retrieves the method number for passing to Invoke.
func MethodAt(obj *IDispatch, method string) (int32, error) {
	var displayID []int32
	displayID, err = obj.GetIDsOfName([]string{method})
	if err != nil {
		return -1, err
	}
	return displayID[0], nil
}

// CallMethod calls method on IDispatch with parameters.
func CallMethod(obj *IDispatch, method string, params ...interface{}) (result *com.Variant, err error) {
	var displayID int32
	displayID, err = MethodAt(obj, method)
	if err != nil {
		return
	}

	if len(params) < 1 {
		result, err = obj.Invoke(displayID, com.MethodDispatchContext)
	} else {
		result, err = obj.Invoke(displayID, com.MethodDispatchContext, params...)
	}

	return
}

// MustCallMethod calls method on IDispatch with parameters or panics.
func MustCallMethod(obj *IDispatch, method string, params ...interface{}) *com.Variant {
	r, err := CallMethod(obj, method, params...)
	if err != nil {
		panic(err.Error())
	}
	return r
}

// GetProperty retrieves property from IDispatch.
func GetProperty(obj *IDispatch, method string, params ...interface{}) (result *com.Variant, err error) {
	var displayID int32
	displayID, err = MethodAt(obj, method)
	if err != nil {
		return
	}

	if len(params) < 1 {
		result, err = obj.Invoke(displayID, com.PropertyGetDispatchContext)
	} else {
		result, err = obj.Invoke(displayID, com.PropertyGetDispatchContext, params...)
	}

	return
}

// MustGetProperty retrieves property from IDispatch or panics.
func MustGetProperty(obj *IDispatch, method string, params ...interface{}) *com.Variant {
	r, err := GetProperty(obj, method, params...)
	if err != nil {
		panic(err.Error())
	}
	return r
}

// PutProperty mutates property.
func PutProperty(obj *IDispatch, method string, params ...interface{}) (result *com.Variant, err error) {
	var displayID int32
	displayID, err = MethodAt(obj, method)
	if err != nil {
		return
	}
	result, err = obj.Invoke(displayID, com.PropertySetDispatchContext, params...)
	return
}

// MustPutProperty mutates property or panics.
func MustPutProperty(obj *IDispatch, method string, params ...interface{}) *com.Variant {
	r, err := PutProperty(obj, method, params...)
	if err != nil {
		panic(err.Error())
	}
	return r
}

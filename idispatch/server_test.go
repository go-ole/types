package idispatch

func ExampleService_iunknown() {
	service, err := Service(&Unknown{}, VirtualTableService(), InterfaceID)

	if err != nil {
		// Aw, something went wrong.
		return
	}
	// Output:
}

func ExampleIUnknownService() {
	service, err := IUnknownService(&Unknown{})

	if err != nil {
		// Aw, something went wrong.
		return
	}
	// Output:
}

# payroll

The following is a package of excel functions that will be useful for speeding up the calculation of payroll. As of now (11/25/2019) there is only the `=GET_PAY_BY_HOUR()` function- which calculates an employee's wages owed based on his times in, times out and hourly rate.

## Installing this package

Currently the functions need to be installed manually by creating modules in VBA and running the code for each function in a seperate module.

## Functions

### `=GET_PAY_BY_HOUR()`

The `=GET_PAY_BY_HOUR()` calculates an employee's wages owed based on his times in, times out and hourly rate. It takes in 3 arguments- the range of an employee's time in, time out and his hourly rate.


<a href='https://github.com/benyamindsmith/payroll/tree/master/'><img src='	Capture_1.PNG' align="center" height="200" /></a>

<a href='https://github.com/benyamindsmith/payroll/tree/master/'><img src='	Capture.PNG' align="center" height="200" /></a>

<a href='https://github.com/benyamindsmith/payroll/tree/master/'><img src='	Capture2.PNG' align="center" height="200" /></a>

Using this function would be equivalent to writing the following formula:

```excel
=(SUM(B2:B4) - SUM(A2:A4)) * 24 * C2
```

This function handles the summation of the total hours and calculates the an employee's wages owed with a simple formula.

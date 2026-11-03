# Base

Originally, VB/VBA did not have an Assert method (now they have Debug.Assert)
But I wanted something more extensive and so I wrote this Assertion class based on [Design by Contract&trade;](https://www.eiffel.org/doc/eiffel/ET-_Design_by_Contract_%28tm%29%2C_Assertions_and_Exceptions), part of the [Eiffel programming language](https://www.eiffel.org/)

These classes provides a basic Assert with additional parameters as well as Pre- and Post-conditions.
It also supports other assertion handlers so you could, for example, write your assertions to a log file, a database or even a web service.

## Getting Started

Import all the classes into your project and then copy the contents of Global.bas into your project and now you have a global Assert object. It's that simple!

### Prerequisites

*(none)*

### Host applications and conditional compilation

The classes compile in Access, Excel and Word. Host-specific code sits behind conditional compilation constants, which you set **per project** in *Tools > Project Properties > Conditional Compilation Arguments* (they do not travel with imported files):

| Constant | Set it in | Effect |
|----------|-----------|--------|
| `InAccess = 1` | Access projects | `cAssertHandler` and `CAssert` use `CurrentDb` to tell open (`.mdb`) from compiled (`.mde`) databases; `cBusyPointer` uses `DoCmd.HourGlass` |
| `InExcel = 1` | Excel projects | `cBusyPointer` uses `Application.Cursor` |
| `InWord = 1` | Word projects | `cBusyPointer` uses `System.Cursor` |

Set several at once as `InAccess = 1 : InExcel = 0`. With none set the code is host-neutral: assertions always offer to debug via a message box, and `cBusyPointer` does nothing.
<!--
### Installing

Import the classes into your project, and either import the `Global` module or copy the contents into your own global module. <br />
That's it!
-->
<!-- 
## Running the tests

Explain how to run the automated tests for this system

### Break down into end to end tests

Explain what these tests test and why

```
Give an example
```

### And coding style tests

Explain what these tests test and why

```
Give an example
```

## Deployment

Add additional notes about how to deploy this on a live system

>> migrations/migration[path].[source]
>> fixtures/fixture[path].[source]

(link to purchase my Access migration tool for $9(?))
-->
<!--
## Built With

* [Dropwizard](http://www.dropwizard.io/1.0.2/docs/) - The web framework used
* [Maven](https://maven.apache.org/) - Dependency Management
* [ROME](https://rometools.github.io/rome/) - Used to generate RSS Feeds
-->
## Components - Index

### Modules

[Globals.bas](#Globals)

### Classes

[CAssert.cls](#CAssert)<br />
[cAssertHandler.cls](#cAssertHandler)<br />
[cSilentAssertionHandler.cls](#cSilentAssertionHandler)

## Components

### Globals
Source: ./Globals.bas

Copy the code from this module into your Global module (or use it as is) to instantiate an Assert object into your project.

#### Usage

`Public Assert As New CAssert`

#### Limitations/Caveats

By default, the cAssert class instantiates a default handler of `cAssertHandler` class.<br />
This class is specific to Access VBA and will display a message if the user is running an uncompiled database.

You can write your won custom assertion handlers by copying the [cSilentAssertionHandler](#cSilentAssertHandler) and modifying it to suit your needs.

## Classes

### cAssert
Source: ./cAssert.cls

#### Properties

`assertionQuit` (Read-only Property) - tells the caller to quit the program<br />
`assertionStop` (Read-only Property) - tells the caller to stop in the code<br />
`assertionContinue` (Read-only Property) - tells the caller to continue code execution

#### Methods

`addHandler(cAssertHandler)` - adds an Assertion Handler<br />
`removeHandler(cAssertHandler)` - removes an Assertion Handler<br />
`Assert(Condition, message [, Location])` - raises an assertion if the Condition is NOT True<br />
`Precondition(Condition, message [, Location])` - a special kind of Assertion (refer to [Design by Contract&trade;](https://www.eiffel.org/doc/eiffel/ET-_Design_by_Contract_%28tm%29%2C_Assertions_and_Exceptions))<br />
`PostCondition(Condition, message [, Location])` - a special kind of Assertion (refer to [Design by Contract&trade;](https://www.eiffel.org/doc/eiffel/ET-_Design_by_Contract_%28tm%29%2C_Assertions_and_Exceptions))

#### Usage

`Public Assert As New CAssert`

#### Limitations/Caveats

### cAssertHandler
Source: ./cAssertHandler.cls

#### Properties

*(none)*

#### Methods

handleAssertion(ByVal message As String) - returns an assertion response type

#### Usage

`Dim ahThis as New cAssertHandler`<br />
`Assert.AddHandler ahThis`<br />
`ahThis.handleAssertion(Your message goes here")` - this is called automagically by the `Assert` methods


#### Limitations/Caveats


### cSilentAssertionHandler
Source: ./cSilentAssertionHandler.cls

#### Properties

`count()` (Read-only property)<br />
`assertions()` (Read-only property)

#### Methods

`reset()` - clears out any unprocessed assertion messages

#### Usage


#### Limitations/Caveats


### cBusyPointer
Source: ./cBusyPointer.cls

#### Properties

_(none)_

#### Methods

`Show` - Displays a busy pointer in place of the default mouse pointer.

#### Usage

`Dim hourGlass As New cBusyPointer`<br />
`hourGlass.Show`

When your object goes out of scope, the Terminate() method will automatically restore your mouse pointer (so you really want to declare this where the work happens)

#### Limitations/Caveats

Due to the way VB/VBA instantiates objects, simply declaring `hourGlass As New cBusyPointer` does nothing.

Instead you need to include an `hourGlass.Show` statement before the busy pointer will appear
e.g. `Dim busyPointer As New cHourglass` doesn't do anything until you do `busyPointer.Show`

## Contributing

We are not currently accepting contributions at this time though you are welcome to clone this 
repository and make your own changes provided you leave all the header blocks and copyright 
notices intact.

## Versioning

We use [SemVer](http://semver.org/) for versioning. For the versions available, see the [tags on this repository](https://github.com/your/project/tags). 

## Authors

* **Paul O'Neill** [pauloneill](https://github.com/pauloneill/)

## License

This project is licensed under the MIT License - see the [LICENSE.md](LICENSE.md) file for details

<!-- ## Acknowledgments -->



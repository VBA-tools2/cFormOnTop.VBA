
# cFormOnTop

Office VBA class to keep UserForms on top of SDI Windows.

This is essentially a republish of Jan Karel Pieterse's article
    <https://jkp-ads.com/articles/keepuserformontop02.asp>.
So all credits go to him!

The main reason for this repository is to bundle all improvements at one point.
Otherwise one has at least to dig through all the comments below the article to
find them.

## Features

- Keep a UserForm on top of SDI Windows[^1] ...

[^1]: SDI stands for "Single Document Interface" which is the new standard
      since Excel 2013. For more information see the
      [Microsoft Docs](https://docs.microsoft.com/en-us/office/vba/excel/concepts/programming-for-the-single-document-interface-in-excel)

## Prerequisites / Dependencies

None.

## How to install / Getting started

Add `cFormOnTop.cls` to your project.
Yes, its that simple.

## Usage / Show it in action

Place

```vba
Private mclsFormOnTop As cFormOnTop

Private Sub UserForm_Initialize()
    Set mclsFormOnTop = New cFormOnTop
    Set mclsFormOnTop.TheUserform = Me
    mclsFormOnTop.InitializeMe
End Sub
```

to the UserForm code. (This is the given example from the source.)

If you want to see it in action, you can also have a look at
`cFormOnTop_demo.xlsm` in the `demo` folder.

## Running Tests

Unfortunately I don't know how to create automated tests/units tests for this
project. If you have an idea, I would love to see it! Please add an issue or
– even better – a pull request (see [Contributing](#contributing)).

But of course one can manually test it. Please have a look at the `tests`
folder.

## Used By

This project is used by (at least) these projects:

- <https://github.com/VBA-tools2/DiffWorksheets>

If you know more, I'll be happy to add them here.

## Known issues and limitations

None that I am aware of.

## Contributing

All contributions are highly welcome!!

If you are new to git/GitHub, please have a look at
    <https://github.com/firstcontributions/first-contributions>
where you will find a lot of useful information for beginners.

I recently was pointed to
    <https://www.conventionalcommits.org>.
which sounds very promising. I'll use them from now on too (and hopefully don't
forget it in a hurry.)

## FAQ

1. What are the `'@...` comments good for in the code?
   You should really have a look at the awesome
   [Rubberduck](https://rubberduckvba.com/) project!

## License

[MIT](https://choosealicense.com/licenses/mit/)

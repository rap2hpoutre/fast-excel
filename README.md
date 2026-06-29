<p align="center">
<img src="https://user-images.githubusercontent.com/36028424/40173202-9a03d68a-5a03-11e8-9968-6b7e3b4f8a1b.png">
</p>

[![Version](https://poser.pugx.org/rap2hpoutre/fast-excel/version?format=flat)](https://packagist.org/packages/rap2hpoutre/fast-excel)
[![License](https://poser.pugx.org/rap2hpoutre/fast-excel/license?format=flat)](https://packagist.org/packages/rap2hpoutre/fast-excel)
[![StyleCI](https://github.styleci.io/repos/128174809/shield?branch=master)](https://github.styleci.io/repos/128174809?branch=master)
[![Tests](https://github.com/rap2hpoutre/fast-excel/actions/workflows/tests.yml/badge.svg)](https://github.com/rap2hpoutre/fast-excel/actions/workflows/tests.yml)
[![Total Downloads](https://poser.pugx.org/rap2hpoutre/fast-excel/downloads)](https://packagist.org/packages/rap2hpoutre/fast-excel)

Fast Excel import/export for Laravel, thanks to [Spout](https://github.com/box/spout).
See [benchmarks](#benchmarks) below.

## Quick start

Install via composer:

```
composer require rap2hpoutre/fast-excel
```

Export a Model to `.xlsx` file:

```php
use Rap2hpoutre\FastExcel\FastExcel;
use App\User;

// Load users
$users = User::all();

// Export all users
(new FastExcel($users))->export('file.xlsx');
```

## Export

Export a Model, Query or **Collection**:

```php
$list = collect([
    [ 'id' => 1, 'name' => 'Jane' ],
    [ 'id' => 2, 'name' => 'John' ],
]);

// export() returns the absolute path to the written file
$path = (new FastExcel($list))->export('file.xlsx');
```

Export `xlsx`, `ods` and `csv`:

```php
$invoices = App\Invoice::orderBy('created_at', 'DESC')->get();
(new FastExcel($invoices))->export('invoices.csv');
```

Export only some attributes specifying columns names:

```php
(new FastExcel(User::all()))->export('users.csv', function ($user) {
    return [
        'Email' => $user->email,
        'First Name' => $user->firstname,
        'Last Name' => strtoupper($user->lastname),
    ];
});
```

Download (from a controller method):

```php
return (new FastExcel(User::all()))->download('file.xlsx');
```

## Import

`import` returns a Collection:

```php
$collection = (new FastExcel)->import('file.xlsx');
```

Import a `csv` with specific delimiter, enclosure characters and "gbk" encoding:

```php
$collection = (new FastExcel)->configureCsv(';', '#', 'gbk')->import('file.csv');
```

Import and insert to database:

```php
$users = (new FastExcel)->import('file.xlsx', function ($line) {
    return User::create([
        'name' => $line['Name'],
        'email' => $line['Email']
    ]);
});
```

## Facades

You may use FastExcel with the optional Facade. Add the following line to ``config/app.php`` under the ``aliases`` key.

````php
'FastExcel' => Rap2hpoutre\FastExcel\Facades\FastExcel::class,
````

Using the Facade, you will not have access to the constructor. You may set your export data using the ``data`` method.

````php
$list = collect([
    [ 'id' => 1, 'name' => 'Jane' ],
    [ 'id' => 2, 'name' => 'John' ],
]);

FastExcel::data($list)->export('file.xlsx');
````

## Global helper

FastExcel provides a convenient global helper to quickly instantiate the FastExcel class anywhere in a Laravel application.

```php
$collection = fastexcel()->import('file.xlsx');
fastexcel($collection)->export('file.xlsx');
```

## Advanced usage

### Export multiple sheets

Export multiple sheets by creating a `SheetCollection`:

```php
$sheets = new SheetCollection([
    User::all(),
    Project::all()
]);
(new FastExcel($sheets))->export('file.xlsx');
```

Use index to specify sheet name:
```php
$sheets = new SheetCollection([
    'Users' => User::all(),
    'Second sheet' => Project::all()
]);
```

### Import multiple sheets

Import multiple sheets by using `importSheets`:

```php
$sheets = (new FastExcel)->importSheets('file.xlsx');
```

You can also import a specific sheet by its number:

```php
$users = (new FastExcel)->sheet(3)->import('file.xlsx');
```

Import multiple sheets with sheets names:

```php
$sheets = (new FastExcel)->withSheetsNames()->importSheets('file.xlsx');
```

### Export large collections with chunk

Export rows one by one to avoid `memory_limit` issues [using `yield`](https://www.php.net/manual/en/language.generators.syntax.php):

```php
function usersGenerator() {
    foreach (User::cursor() as $user) {
        yield $user;
    }
}

// Export consumes only a few MB, even with 10M+ rows.
(new FastExcel(usersGenerator()))->export('test.xlsx');
```

### Import large files (low memory)

`import` returns a Collection containing every row, so memory grows with the size of
the file. On a file larger than your PHP `memory_limit`, the default `import()` fails
outright with `Allowed memory size of N bytes exhausted`. To import a large file
without running out of memory, pass a callback and **return `null`** — each row is
then processed but not accumulated, so memory stays flat:

```php
// Memory stays flat regardless of the number of rows.
(new FastExcel)->import('file.xlsx', function ($line) {
    User::create([
        'name'  => $line['Name'],
        'email' => $line['Email'],
    ]);

    return null; // don't keep the row in memory
});
```

> If the callback returns a value (for example the created model), that value is
> collected and returned to you — handy for small files, but it keeps every row in
> memory. Return `null` when you only need the side effect (e.g. inserting rows).

Importing a 730,000-row file (8 columns), measured under a constrained
`memory_limit`:

| How you import | Peak memory | Result |
| --- | --- | --- |
| `import($file)` (returns a Collection) | ~440 MB | **fails** once it exceeds `memory_limit` |
| `import($file, fn ($row) => null)` (streaming) | ~4 MB | always completes |

Reading speed is the same either way — it is dominated by the underlying
[OpenSpout](https://github.com/openspout/openspout) parser, not by how the rows are
returned. The difference above is memory, which is what lets very large files finish
at all.

### Add header and rows style

Add header and rows style with `headerStyle` and `rowsStyle` methods.

```php
use OpenSpout\Common\Entity\Style\Style;

$header_style = (new Style())->setFontBold();

$rows_style = (new Style())
    ->setFontSize(15)
    ->setShouldWrapText()
    ->setBackgroundColor("EDEDED");

return (new FastExcel($list))
    ->headerStyle($header_style)
    ->rowsStyle($rows_style)
    ->download('file.xlsx');
```

You can also style each header column individually with `setHeaderColumnStyles`
(the header-row counterpart of `setColumnStyles`). Keys are the zero-based
column positions:

```php
use OpenSpout\Common\Entity\Style\Color;
use OpenSpout\Common\Entity\Style\Style;

return (new FastExcel($list))
    ->setHeaderColumnStyles([
        0 => (new Style())->setBackgroundColor(Color::YELLOW),
        1 => (new Style())->setFontColor(Color::BLUE),
    ])
    ->download('file.xlsx');
```

### Export values as strings or numbers

By default numbers are written as numbers and strings as strings. Use
`stringValues()` to force every value to a text cell — handy to keep leading
zeros or long numeric IDs (e.g. phone numbers) intact:

```php
// 0660123456 stays "0660123456" instead of becoming 660123456
(new FastExcel($users))->stringValues()->export('users.xlsx');
```

Need finer control? `setColumnFormat()` overrides the type per column and takes
precedence over `stringValues()`:

```php
(new FastExcel($users))
    ->stringValues()                                  // text by default
    ->setColumnFormat([
        'id'    => 'number',                          // keep as number
        'phone' => 'string',                          // keep as text
    ])
    ->export('users.xlsx');
```

## Why?

FastExcel is intended at being Laravel-flavoured [Spout](https://github.com/box/spout):
a simple, but elegant wrapper around [Spout](https://github.com/box/spout) with the goal
of simplifying **imports and exports**. It could be considered as a faster (and memory friendly) alternative
to [Laravel Excel](https://laravel-excel.com/), with less features.
Use it only for simple tasks.

## Benchmarks

> Tested on a MacBook Pro 2015 2,7 GHz Intel Core i5 16 Go 1867 MHz DDR3.
Testing a XLSX export for 10000 lines, 20 columns with random data, 10 iterations, 2018-04-05. **Don't trust benchmarks.**

|   | Average memory peak usage  | Execution time |
|---|---|---|
| Laravel Excel  | 123.56 M  | 11.56 s |
| FastExcel  | 2.09 M | 2.76 s |

Still, remember that [Laravel Excel](https://laravel-excel.com/) **has many more features.**

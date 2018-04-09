# Laravel Fast Excel

[![Packagist](https://img.shields.io/packagist/v/rap2hpoutre/fast-excel.svg)]()
[![Packagist](https://img.shields.io/packagist/l/rap2hpoutre/fast-excel.svg)](https://packagist.org/packages/rap2hpoutre/fast-excel)
[![Build Status](https://travis-ci.org/rap2hpoutre/fast-excel.svg?branch=master)](https://travis-ci.org/rap2hpoutre/fast-excel)
[![Scrutinizer Code Quality](https://scrutinizer-ci.com/g/rap2hpoutre/fast-excel/badges/quality-score.png?b=master)](https://scrutinizer-ci.com/g/rap2hpoutre/fast-excel/?branch=master)

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

## Usage

Use Fast Excel to import and export Excel files.

### Export

You can export a Model or a **Collection**:

```php
$list = collect([
    [ 'id' => 1, 'name' => 'Jane' ],
    [ 'id' => 2, 'name' => 'John' ],
]);

(new FastExcel($list))->export('file.xlsx');
```

You can export `xlsx`, `ods` and `csv`:

```php
$invoices = App\Invoice::orderBy('created_at', 'DESC')->get();

(new FastExcel($invoices))->export('invoices.csv');
```

Export only some attributes and choose columns name:

```php
(new FastExcel(User::all()))->export('users.csv', function ($user) {
    return [
        'Email' => $user->email,
        'First Name' => $user->firstname,
        'Last Name' => strtoupper($user->lastname),
    ];
});
```

### Import

`import` returns a Collection:

```php
$collection = (new FastExcel)->import('file.xlsx');
```

Import a `csv` with a specific delimiter and enclosure characters.

```php
$collection = (new FastExcel)->configureCsv(';', '#')->import('file.csv');
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

## Why?

Laravel Fast Excel is intended at being Laravel-flavoured [Spout](https://github.com/box/spout): 
a simple, but elegant wrapper around [Spout](https://github.com/box/spout) with the goal 
of simplifying **imports and exports**. 

It could be considered as a faster (and memory friendly) alternative 
to [Laravel Excel](https://laravel-excel.maatwebsite.nl/), with **many less** features. 
Use it only for very simple tasks.

## Benchmarks

> Tested on a MacBook Pro 2015 2,7 GHz Intel Core i5 16 Go 1867 MHz DDR3. 
Testing a XLSX export for 10000 lines, 20 columns with random data, 10 iterations, 2018-04-05. **Don't trust benchmarks.**



|   | Average memory peak usage  | Execution time |
|---|---|---|
| Laravel Excel  | 123.56 M  | 11.56 s |
| Laravel Fast Excel  | 2.09 M | 2.76 s |

Still, remember that [Laravel Excel](https://laravel-excel.maatwebsite.nl/) **has many more feature.**
Please help me improve benchmarks, more tests re coming. Feel free to criticize.
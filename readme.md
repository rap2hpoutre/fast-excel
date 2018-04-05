# Laravel Stripe Connect

[![Packagist](https://img.shields.io/packagist/v/rap2hpoutre/laravel-stripe-connect.svg)]()
[![Packagist](https://img.shields.io/packagist/l/rap2hpoutre/laravel-stripe-connect.svg)](https://packagist.org/packages/rap2hpoutre/laravel-stripe-connect)
[![Build Status](https://travis-ci.org/rap2hpoutre/laravel-stripe-connect.svg?branch=master)](https://travis-ci.org/rap2hpoutre/laravel-stripe-connect)
[![Scrutinizer Code Quality](https://scrutinizer-ci.com/g/rap2hpoutre/laravel-stripe-connect/badges/quality-score.png?b=master)](https://scrutinizer-ci.com/g/rap2hpoutre/laravel-stripe-connect/?branch=master)

> Marketplaces and platforms use Stripe Connect to accept money and pay out to third parties. Connect provides a complete set of building blocks to support virtually any business model, including on-demand businesses, e‑commerce, crowdfunding, fintech, and travel and events. 

Create a marketplace application with this helper for [Stripe Connect](https://stripe.com/connect).

## Installation

Install via composer

```
composer require rap2hpoutre/laravel-stripe-connect
```

Add your stripe credentials in `.env`:

```
STRIPE_KEY=pk_test_XxxXXxXXX
STRIPE_SECRET=sk_test_XxxXXxXXX
```

Run migrations:

```
php artisan migrate
```

## Usage

You can make a single payment from a user to another user
 or save a customer card for later use. Just remember to
 import the base class via:
 
```php
use Rap2hpoutre\LaravelStripeConnect\StripeConnect;
```

### Example #1: direct charge

The customer gives his credentials via Stripe Checkout and is charged.
It's a one shot process. `$customer` and `$vendor` must be `User` instances. The `$token` must have been created using [Checkout](https://stripe.com/docs/checkout/tutorial) or [Elements](https://stripe.com/docs/stripe-js).

```php
StripeConnect::transaction($token)
    ->amount(1000, 'usd')
    ->from($customer)
    ->to($vendor)
    ->create(); 
```

### Example #2: save a customer then charge later

Sometimes, you may want to register a card then charge later.
First, create the customer.

```php
StripeConnect::createCustomer($token, $customer);
```

Then, (later) charge the customer without token.

```php
StripeConnect::transaction()
    ->amount(1000, 'usd')
    ->useSavedCustomer()
    ->from($customer)
    ->to($vendor)
    ->create(); 
```

### Exemple #3: create a vendor account

You may want to create the vendor account before charging anybody.
Just call `createAccount` with a `User` instance.

```php
StripeConnect::createAccount($vendor);
```

### Exemple #4: Charge with application fee

```php
StripeConnect::transaction()
    ->amount(1000, 'usd')
    ->fee(50)
    ->from($customer)
    ->to($vendor)
    ->create(); 
```

<?php

namespace Rap2hpoutre\FastExcel\Tests;

use Illuminate\Database\Eloquent\Model;

/**
 * Class User.
 *
 * Minimal Eloquent model used to test exporting/importing model collections.
 */
class User extends Model
{
    public $timestamps = false;

    protected $fillable = ['name', 'email'];
}

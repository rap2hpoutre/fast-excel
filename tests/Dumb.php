<?php

namespace Rap2hpoutre\FastExcel\Tests;

use Illuminate\Database\Eloquent\Model;


/**
 * Class Dumb
 * @package Rap2hpoutre\FastExcel\Tests
 */
class Dumb extends Model
{
    public $data;

    /**
     * Dumb constructor.
     * @param $data
     */
    public function __construct($data)
    {
        parent::__construct();
        $this->data = $data;
    }
}
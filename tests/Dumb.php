<?php

namespace Rap2hpoutre\FastExcel\Tests;

use Illuminate\Database\Eloquent\Model;

/**
 * Class Dumb.
 */
class Dumb extends Model
{
    public $data;

    /**
     * Dumb constructor.
     *
     * @param $data
     */
    public function __construct(string $data = '')
    {
        parent::__construct();
        $this->data = $data;
    }
}

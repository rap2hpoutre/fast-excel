<?php

namespace Rap2hpoutre\FastExcel\Tests;

use Illuminate\Database\Capsule\Manager as Capsule;
use OpenSpout\Common\Exception\IOException;
use Rap2hpoutre\FastExcel\FastExcel;

/**
 * Class EloquentModelTest.
 */
class EloquentModelTest extends TestCase
{
    /**
     * Issue #12: export and re-import a real Eloquent model collection
     * (backed by an in-memory SQLite database), so model attributes round-trip
     * through the writer/reader and not just plain collections.
     *
     * @throws IOException
     * @throws \OpenSpout\Writer\Exception\WriterNotOpenedException
     * @throws \OpenSpout\Reader\Exception\ReaderNotOpenedException
     * @throws \OpenSpout\Common\Exception\UnsupportedTypeException
     */
    public function testExportAndImportEloquentModels()
    {
        $capsule = new Capsule();
        $capsule->addConnection(['driver' => 'sqlite', 'database' => ':memory:']);
        $capsule->setAsGlobal();
        $capsule->bootEloquent();

        $capsule->schema()->create('users', function ($table) {
            $table->increments('id');
            $table->string('name');
            $table->string('email');
        });

        $user = new User();
        $user->newQuery()->insert([
            ['name' => 'Joe', 'email' => 'joe@example.com'],
            ['name' => 'Jane', 'email' => 'jane@example.com'],
        ]);

        $file = __DIR__.'/issue_12.xlsx';
        (new FastExcel($user->newQuery()->get()))->export($file);

        $imported = (new FastExcel())->import($file);

        $this->assertCount(2, $imported);
        $this->assertEquals(['id', 'name', 'email'], array_keys($imported[0]));
        $this->assertEquals('Joe', $imported[0]['name']);
        $this->assertEquals('joe@example.com', $imported[0]['email']);
        $this->assertEquals('Jane', $imported[1]['name']);

        $capsule->schema()->drop('users');
        unlink($file);
    }
}

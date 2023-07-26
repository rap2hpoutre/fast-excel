<?php

namespace Smart145\FastExcel\Providers;

use Illuminate\Support\ServiceProvider;

class FastExcelServiceProvider extends ServiceProvider
{
    /**
     * Bootstrap any application services.
     *
     * @return void
     */
    public function boot()
    {
        //
    }

    /**
     * Register any application services.
     *
     * @SuppressWarnings("unused")
     *
     * @return void
     */
    public function register()
    {
        $this->app->bind('fastexcel', function ($app, $data = null) {
            if (is_array($data)) {
                $data = collect($data);
            }

            return new \Smart145\FastExcel\FastExcel($data);
        });
    }
}

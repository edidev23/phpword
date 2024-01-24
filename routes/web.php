<?php

use Illuminate\Support\Facades\Route;

/*
|--------------------------------------------------------------------------
| Web Routes
|--------------------------------------------------------------------------
|
| Here is where you can register web routes for your application. These
| routes are loaded by the RouteServiceProvider within a group which
| contains the "web" middleware group. Now create something great!
|
*/

Route::get('/', function () {
    return view('welcome');
});

Route::get('/create-word', [App\Http\Controllers\PhpwordController::class, 'createWord']);
Route::get('/modify-word', [App\Http\Controllers\PhpwordController::class, 'modifyDocument']);
Route::get('/generate-word', [App\Http\Controllers\PhpwordController::class, 'generateWord']);

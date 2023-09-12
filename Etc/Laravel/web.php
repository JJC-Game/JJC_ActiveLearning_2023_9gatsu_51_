<?php

use Illuminate\Support\Facades\Route;
use App\Http\Controllers\AppUserController;

/*
|--------------------------------------------------------------------------
| Web Routes
|--------------------------------------------------------------------------
|
| Here is where you can register web routes for your application. These
| routes are loaded by the RouteServiceProvider and all of them will
| be assigned to the "web" middleware group. Make something great!
|
*/

Route::get('/app_user/{id}', [AppUserController::class, 'app_user'])->name('al.app_user');
Route::get('/play_chara_gacha/{id}', [AppUserController::class, 'play_chara_gacha'])->name('al.play_chara_gacha');

Route::get('/', function () {
    return view('welcome');
});

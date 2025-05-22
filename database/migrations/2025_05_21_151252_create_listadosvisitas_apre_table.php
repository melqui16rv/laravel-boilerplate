<?php

use Illuminate\Database\Migrations\Migration;
use Illuminate\Database\Schema\Blueprint;
use Illuminate\Support\Facades\Schema;

return new class extends Migration
{
    /**
     * Run the migrations.
     */
    public function up(): void
    {
        Schema::disableForeignKeyConstraints();

        Schema::create('listadosvisitas_apre', function (Blueprint $table) {
            $table->integer('id_visita')->primary();
            $table->string('nodo', 100)->nullable()->default('Cundinamarca');
            $table->string('encargado', 155);
            $table->integer('numAsistentes');
            $table->dateTime('fechaCharla');
            $table->dateTime('fecha_insert')->useCurrent();
        });

        Schema::enableForeignKeyConstraints();
    }

    /**
     * Reverse the migrations.
     */
    public function down(): void
    {
        Schema::dropIfExists('listadosvisitas_apre');
    }
};

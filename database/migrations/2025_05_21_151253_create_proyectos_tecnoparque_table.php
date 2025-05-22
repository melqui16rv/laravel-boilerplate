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

        Schema::create('proyectos_tecnoparque', function (Blueprint $table) {
            $table->integer('id_PBT')->primary();
            $table->enum('tipo', [""]);
            $table->string('nombre_linea', 55);
            $table->integer('terminados')->nullable()->default('DEFAULT NULL');
            $table->integer('en_proceso')->nullable()->default('DEFAULT NULL');
            $table->dateTime('fecha_actualizacion')->useCurrent();
        });

        Schema::enableForeignKeyConstraints();
    }

    /**
     * Reverse the migrations.
     */
    public function down(): void
    {
        Schema::dropIfExists('proyectos_tecnoparque');
    }
};

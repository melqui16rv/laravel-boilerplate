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

        Schema::create('usuario', function (Blueprint $table) {
            $table->string('numero_documento', 79)->primary();
            $table->string('tipo_doc', 100);
            $table->string('nombre_completo', 300)->nullable()->default('DEFAULT NULL');
            $table->string('contrase\0f1a', 200)->nullable()->default('DEFAULT NULL');
            $table->string('email', 200);
            $table->string('telefono', 50);
            $table->string('id_rol', 10);
        });

        Schema::enableForeignKeyConstraints();
    }

    /**
     * Reverse the migrations.
     */
    public function down(): void
    {
        Schema::dropIfExists('usuario');
    }
};

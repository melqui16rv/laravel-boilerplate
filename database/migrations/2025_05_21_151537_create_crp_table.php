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
        Schema::create('crp', function (Blueprint $table) {
            $table->string('CODIGO_CRP', 55)->primary();
            $table->string('CODIGO_CDP', 55);
            $table->string('Numero_Documento', 55)->nullable()->default('DEFAULT NULL');
            $table->date('Fecha_de_Registro')->nullable()->default('DEFAULT NULL');
            $table->dateTime('Fecha_de_Creacion')->nullable()->default('DEFAULT NULL');
            $table->string('Estado', 255)->nullable()->default('DEFAULT NULL');
            $table->string('Dependencia', 255)->nullable()->default('DEFAULT NULL');
            $table->text('Rubro')->nullable()->default('DEFAULT NULL');
            $table->text('Descripcion')->nullable()->default('DEFAULT NULL');
            $table->string('Fuente', 100)->nullable()->default('DEFAULT NULL');
            $table->decimal('Valor_Inicial', 15, 2)->nullable()->default('DEFAULT NULL');
            $table->decimal('Valor_Operaciones', 15, 2)->nullable()->default('DEFAULT NULL');
            $table->decimal('Valor_Actual', 15, 2)->nullable()->default('DEFAULT NULL');
            $table->decimal('Saldo_por_Utilizar', 15, 2)->nullable()->default('DEFAULT NULL');
            $table->string('Tipo_Identificacion', 255)->nullable()->default('DEFAULT NULL');
            $table->string('Identificacion', 255)->nullable()->default('DEFAULT NULL');
            $table->string('Nombre_Razon_Social', 255)->nullable()->default('DEFAULT NULL');
            $table->string('Medio_de_Pago', 255)->nullable()->default('DEFAULT NULL');
            $table->string('Tipo_Cuenta', 255)->nullable()->default('DEFAULT NULL');
            $table->string('Numero_Cuenta', 255)->nullable()->default('DEFAULT NULL');
            $table->string('Estado_Cuenta', 255)->nullable()->default('DEFAULT NULL');
            $table->string('Entidad_Nit', 255)->nullable()->default('DEFAULT NULL');
            $table->text('Entidad_Descripcion')->nullable()->default('DEFAULT NULL');
            $table->string('Solicitud_CDP', 55)->nullable()->default('DEFAULT NULL');
            $table->string('CDP', 55)->nullable()->default('DEFAULT NULL');
            $table->string('Compromisos', 55)->nullable()->default('DEFAULT NULL');
            $table->text('Cuentas_por_Pagar')->nullable()->default('DEFAULT NULL');
            $table->text('Obligaciones')->nullable()->default('DEFAULT NULL');
            $table->text('Ordenes_de_Pago')->nullable()->default('DEFAULT NULL');
            $table->decimal('Reintegros', 15, 2)->nullable()->default('DEFAULT NULL');
            $table->dateTime('Fecha_Documento_Soporte')->nullable()->default('DEFAULT NULL');
            $table->string('Tipo_Documento_Soporte', 255)->nullable()->default('DEFAULT NULL');
            $table->string('Numero_Documento_Soporte', 255)->nullable()->default('DEFAULT NULL');
            $table->text('Observaciones')->nullable()->default('DEFAULT NULL');
        });
    }

    /**
     * Reverse the migrations.
     */
    public function down(): void
    {
        Schema::dropIfExists('crp');
    }
};

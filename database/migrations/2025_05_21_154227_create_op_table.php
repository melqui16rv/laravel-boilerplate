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
        Schema::create('op', function (Blueprint $table) {
            $table->string('CODIGO_OP', 55)->primary();
            $table->string('CODIGO_CRP', 55);
            $table->string('CODIGO_CDP', 55);
            $table->string('Numero_Documento', 55)->nullable()->default('DEFAULT NULL');
            $table->date('Fecha_de_Registro')->nullable()->default('DEFAULT NULL');
            $table->dateTime('Fecha_de_Pago')->nullable()->default('DEFAULT NULL');
            $table->string('Estado', 255)->nullable()->default('DEFAULT NULL');
            $table->decimal('Valor_Bruto', 15, 2)->nullable()->default('DEFAULT NULL');
            $table->decimal('Valor_Deducciones', 15, 2)->nullable()->default('DEFAULT NULL');
            $table->decimal('Valor_Neto', 15, 2)->nullable()->default('DEFAULT NULL');
            $table->string('Tipo_Beneficiario', 155)->nullable()->default('DEFAULT NULL');
            $table->string('Vigencia_Presupuestal', 155)->nullable()->default('DEFAULT NULL');
            $table->string('Tipo_Identificacion', 255)->nullable()->default('DEFAULT NULL');
            $table->string('Identificacion', 255)->nullable()->default('DEFAULT NULL');
            $table->string('Nombre_Razon_Social', 255)->nullable()->default('DEFAULT NULL');
            $table->string('Medio_de_Pago', 255)->nullable()->default('DEFAULT NULL');
            $table->string('Tipo_Cuenta', 255)->nullable()->default('DEFAULT NULL');
            $table->string('Numero_Cuenta', 255)->nullable()->default('DEFAULT NULL');
            $table->string('Estado_Cuenta', 255)->nullable()->default('DEFAULT NULL');
            $table->string('Entidad_Nit', 255)->nullable()->default('DEFAULT NULL');
            $table->text('Entidad_Descripcion')->nullable()->default('DEFAULT NULL');
            $table->string('Dependencia', 255)->nullable()->default('DEFAULT NULL');
            $table->text('Dependencia_Descripcion')->nullable()->default('DEFAULT NULL');
            $table->text('Rubro')->nullable()->default('DEFAULT NULL');
            $table->text('Descripcion')->nullable()->default('DEFAULT NULL');
            $table->string('Fuente', 100)->nullable()->default('DEFAULT NULL');
            $table->string('Recurso', 255)->nullable()->default('DEFAULT NULL');
            $table->string('Sit', 155)->nullable()->default('DEFAULT NULL');
            $table->decimal('Valor_Pesos', 15, 2)->nullable()->default('DEFAULT NULL');
            $table->decimal('Valor_Moneda', 15, 2)->nullable()->default('DEFAULT NULL');
            $table->decimal('Valor_Reintegrado_Pesos', 15, 2)->nullable()->default('DEFAULT NULL');
            $table->decimal('Valor_Reintegrado_Moneda', 15, 2)->nullable()->default('DEFAULT NULL');
            $table->string('Tesoreria_Pagadora', 100)->nullable()->default('DEFAULT NULL');
            $table->string('Identificacion_Pagaduria', 255)->nullable()->default('DEFAULT NULL');
            $table->string('Cuenta_Pagaduria', 255)->nullable()->default('DEFAULT NULL');
            $table->string('Endosada', 55)->nullable()->default('DEFAULT NULL');
            $table->string('Tipo_Identificacion2', 255)->nullable()->default('DEFAULT NULL');
            $table->string('Identificacion3', 255)->nullable()->default('DEFAULT NULL');
            $table->string('Razon_social', 255)->nullable()->default('DEFAULT NULL');
            $table->string('Numero_Cuenta4', 255)->nullable()->default('DEFAULT NULL');
            $table->text('Concepto_Pago')->nullable()->default('DEFAULT NULL');
            $table->string('Solicitud_CDP', 55)->nullable()->default('DEFAULT NULL');
            $table->string('CDP', 55)->nullable()->default('DEFAULT NULL');
            $table->string('Compromisos', 55)->nullable()->default('DEFAULT NULL');
            $table->text('Cuentas_por_Pagar')->nullable()->default('DEFAULT NULL');
            $table->date('Fecha_Cuentas_por_Pagar')->nullable()->default('DEFAULT NULL');
            $table->text('Obligaciones')->nullable()->default('DEFAULT NULL');
            $table->text('Ordenes_de_Pago')->nullable()->default('DEFAULT NULL');
            $table->decimal('Reintegros', 15, 2)->nullable()->default('DEFAULT NULL');
            $table->date('Fecha_Doc_Soporte_Compromiso')->nullable()->default('DEFAULT NULL');
            $table->string('Tipo_Doc_Soporte_Compromiso', 100)->nullable()->default('DEFAULT NULL');
            $table->string('Num_Doc_Soporte_Compromiso', 100)->nullable()->default('DEFAULT NULL');
            $table->text('Objeto_del_Compromiso')->nullable()->default('DEFAULT NULL');
        });
    }

    /**
     * Reverse the migrations.
     */
    public function down(): void
    {
        Schema::dropIfExists('op');
    }
};

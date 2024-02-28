<template>
  <div>
    <h2>Formulario de Datos</h2>
    <div class="formulario">
      <!-- Campos del formulario -->
      <div class="input-group">
        <label for="nombre_epp">Nombre del EPP:</label>
        <input type="text" id="nombre_epp" v-model="datos.nombre_epp">
      </div>       
      <div class="input-group">
        <label for="parte_cuerpo_proteger">Parte del Cuerpo a Proteger:</label>
        <input type="text" id="parte_cuerpo_proteger" v-model="datos.parte_cuerpo_proteger">
      </div>
      <div class="input-group">
        <label for="riesgo_controlado">Riesgo Controlado:</label>
        <input type="text" id="riesgo_controlado" v-model="datos.riesgo_controlado">
      </div>
      <div class="input-group">
        <label for="cargo_asociado">Cargo Asociado:</label>
        <input type="text" id="cargo_asociado" v-model="datos.cargo_asociado">
      </div>
      <div class="input-group">
        <label for="especificacion_tecnica">Especificación Técnica:</label>
        <input type="text" id="especificacion_tecnica" v-model="datos.especificacion_tecnica">
      </div>
      <div class="input-group">
        <label for="uso">Uso:</label>
        <input type="text" id="uso" v-model="datos.uso">
      </div>
      <div class="input-group">
        <label for="mantenimiento">Mantenimiento:</label>
        <input type="text" id="mantenimiento" v-model="datos.mantenimiento">
      </div>
      <div class="input-group">
        <label for="vida_util">Vida Útil:</label>
        <input type="text" id="vida_util" v-model="datos.vida_util">
      </div>
      <div class="input-group">
        <label for="reposicion">Reposición:</label>
        <input type="text" id="reposicion" v-model="datos.reposicion">
      </div>
      <div class="input-group">
        <label for="disposicion_final">Disposición Final:</label>
        <input type="text" id="disposicion_final" v-model="datos.disposicion_final">
      </div>
    </div>

    <!-- Botones para agregar item y exportar a Excel -->
    <button @click="agregarItem">Agregar Item</button>
    <button @click="exportarExcel">Exportar a Excel</button>

    <!-- Tabla de datos ingresados -->
    <h3>Tabla de Datos Ingresados</h3>
    <div class="tabla-container">
      <table class="tabla-datos">
        <thead>
          <tr>
            <th></th>
            <th></th>
            <th></th>
            <th></th>
            <th></th>
            <th></th>
            <th></th>
            <th></th>
            <th></th>
            <th></th>
          </tr>
        </thead>
        <tbody>
          <tr v-for="(item, index) in listaDatos" :key="index">
            <td>{{ item.nombre_epp }}</td>
            <td>{{ item.parte_cuerpo_proteger }}</td>
            <td>{{ item.riesgo_controlado }}</td>
            <td>{{ item.cargo_asociado }}</td>
            <td>{{ item.especificacion_tecnica }}</td>
            <td>{{ item.uso }}</td>
            <td>{{ item.mantenimiento }}</td>
            <td>{{ item.vida_util }}</td>
            <td>{{ item.reposicion }}</td>
            <td>{{ item.disposicion_final }}</td>
          </tr>
        </tbody>
      </table>
    </div>
  </div>
</template>

<script>
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';

export default {
  name: 'FormularioView',
  data() {
    return {
      datos: {
        nombre_epp: '',
        parte_cuerpo_proteger: '',
        riesgo_controlado: '',
        cargo_asociado: '',
        especificacion_tecnica: '',
        uso: '',
        mantenimiento: '',
        vida_util: '',
        reposicion: '',
        disposicion_final: ''
      },
      listaDatos: []
    };
  },
  methods: {
    agregarItem() {
      this.listaDatos.push({ ...this.datos });
      this.resetForm();
    },
    resetForm() {
      this.datos = {
        nombre_epp: '',
        parte_cuerpo_proteger: '',
        riesgo_controlado: '',
        cargo_asociado: '',
        especificacion_tecnica: '',
        uso: '',
        mantenimiento: '',
        vida_util: '',
        reposicion: '',
        disposicion_final: ''
      };
    },
    exportarExcel() {
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet('Datos');

      // Encabezados
      const headers = [
        'Nombre del EPP',
        'Parte del Cuerpo a Proteger',
        'Riesgo Controlado',
        'Cargo Asociado',
        'Especificación Técnica',
        'Uso',
        'Mantenimiento',
        'Vida Útil',
        'Reposición',
        'Disposición Final'
      ];
      worksheet.mergeCells('A1:J1');
      const titleCell = worksheet.getCell('A1');
      titleCell.value = 'MATRIZ DE IDENTIFICACIÓN DE ELEMENTOS DE PROTECCIÓN PERSONAL';
      titleCell.font = { bold: true };
      titleCell.alignment = { vertical: 'middle', horizontal: 'center' };
      titleCell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '8F8F8F' }
      };
      const headerRow = worksheet.addRow(headers);
      headerRow.font = { bold: true };

      headerRow.eachCell((cell) => {
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
      });

      worksheet.getRow(1).height = 39;
      worksheet.getRow(2).height = 39;

    
      this.listaDatos.forEach(item => {
        const row = worksheet.addRow(Object.values(item));
        row.eachCell({ includeEmpty: true }, (cell) => {
          cell.alignment = { vertical: 'middle', horizontal: 'center' }; 
          cell.alignment.wrapText = true; 
        });
      });

      worksheet.eachRow((row) => {
        row.eachCell((cell) => {
          cell.border = {
            top: { style: 'thin', color: { argb: '000000' } },
            left: { style: 'thin', color: { argb: '000000' } },
            bottom: { style: 'thin', color: { argb: '000000' } },
            right: { style: 'thin', color: { argb: '000000' } },
          };
        });
      });

      worksheet.getColumn('A').width = 24;
      worksheet.getColumn('B').width = 27;
      worksheet.getColumn('C').width = 26;
      worksheet.getColumn('D').width = 35;
      worksheet.getColumn('E').width = 79;
      worksheet.getColumn('F').width = 51;
      worksheet.getColumn('G').width = 44;
      worksheet.getColumn('H').width = 35;
      worksheet.getColumn('I').width = 35;
      worksheet.getColumn('J').width = 35;

      workbook.xlsx.writeBuffer().then(buffer => {
        const blob = new Blob([buffer]);
        const fileName = 'datos.xlsx';
        saveAs(blob, fileName);
      });
    }
  }
};
</script>

<style>
.formulario {
  margin-bottom: 20px;
}

.input-group {
  margin-bottom: 10px;
}

.tabla-container {
  border: 2px solid #ccc;
  border-radius: 5px;
  padding: 10px;
  margin-top: 20px;
}

.tabla-datos {
  width: 100%;
  border-collapse: collapse;
}

.tabla-datos th, .tabla-datos td {
  border: 1px solid #000000;
  padding: 8px;
  text-align: center;
}

.tabla-datos th {
  background-color: #f2f2f2;
}

.tabla-datos tr:nth-child(even) {
  background-color: #f2f2f2;
}

.tabla-datos tr:hover {
  background-color: #ddd; 
}
</style>

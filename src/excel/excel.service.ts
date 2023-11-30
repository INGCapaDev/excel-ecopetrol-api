import { Injectable } from '@nestjs/common';
import { randomUUID } from 'crypto';
import * as ExcelJS from 'exceljs';

// cSpell: disable
@Injectable()
export class ExcelService {
  private readonly TEMPLATE_EXCEL_PATH: string;
  private readonly EMPTY_CELL = '-';
  constructor() {
    this.TEMPLATE_EXCEL_PATH = './public/templates/template-excel.xlsx';
  }

  validateBody(body: { data: Data }) {
    if (!body.data) throw new Error('No se recibieron datos');
  }

  async createExcel(data: Data) {
    const template = this.TEMPLATE_EXCEL_PATH.toString();
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(template);
    const worksheet = workbook.getWorksheet(1);

    let row = 6;
    data.forEach((item) => {
      worksheet.getCell(`A${row}`).value = item.pozo || this.EMPTY_CELL;
      worksheet.getCell(`B${row}`).value = item.fecha || this.EMPTY_CELL;
      worksheet.getCell(`C${row}`).value = item.marcaVSD || this.EMPTY_CELL;
      worksheet.getCell(`D${row}`).value = item.vsdPotencia || this.EMPTY_CELL;
      worksheet.getCell(`E${row}`).value = item.vsdCorriente || this.EMPTY_CELL;
      worksheet.getCell(`F${row}`).value =
        item.vsdEstadoSistemaVentilacion || this.EMPTY_CELL;
      worksheet.getCell(`G${row}`).value =
        item.comunicacionSCADA || this.EMPTY_CELL;
      worksheet.getCell(`H${row}`).value = item.marcaMotor || this.EMPTY_CELL;
      worksheet.getCell(`I${row}`).value =
        item.motorPotencia || this.EMPTY_CELL;
      worksheet.getCell(`J${row}`).value =
        item.motorCorriente || this.EMPTY_CELL;

      worksheet.getCell(`K${row}`).value = item.numeroPolos || this.EMPTY_CELL;
      worksheet.getCell(`L${row}`).value =
        item.motorVelocidadRPM || this.EMPTY_CELL;
      worksheet.getCell(`M${row}`).value = item.marcaCabezal || this.EMPTY_CELL;
      worksheet.getCell(`N${row}`).value =
        item.sistemaLubricacion || this.EMPTY_CELL;
      worksheet.getCell(`O${row}`).value =
        item.modeloCabezal || this.EMPTY_CELL;
      worksheet.getCell(`P${row}`).value = item.sNCabezal || this.EMPTY_CELL;
      worksheet.getCell(`Q${row}`).value =
        item.relacionPoleas || this.EMPTY_CELL;
      worksheet.getCell(`R${row}`).value =
        item.controladorSAM || this.EMPTY_CELL;
      worksheet.getCell(`S${row}`).value = item.tipo || this.EMPTY_CELL;
      worksheet.getCell(`T${row}`).value = item.marca || this.EMPTY_CELL;
      worksheet.getCell(`U${row}`).value = item.motorotador || this.EMPTY_CELL;
      worksheet.getCell(`V${row}`).value =
        item.rotadorSobreRejillasContraPozo || this.EMPTY_CELL;
      worksheet.getCell(`W${row}`).value = item.RPM || this.EMPTY_CELL;
      worksheet.getCell(`X${row}`).value =
        item.corrienteMotor || this.EMPTY_CELL;
      worksheet.getCell(`Y${row}`).value =
        item.torqueBarraLisa || this.EMPTY_CELL;
      worksheet.getCell(`Z${row}`).value =
        item.torqueMotorPorcentaje || this.EMPTY_CELL;
      worksheet.getCell(`AA${row}`).value =
        item.presionLineaProduccionCabezaPozo || this.EMPTY_CELL;
      worksheet.getCell(`AB${row}`).value =
        item.presionLineaGasCabezaPozo || this.EMPTY_CELL;
      worksheet.getCell(`AC${row}`).value =
        item.cajaEmpaques || this.EMPTY_CELL;
      worksheet.getCell(`AD${row}`).value =
        item.inyeccionQuimico || this.EMPTY_CELL;
      worksheet.getCell(`AE${row}`).value = item.contraPozo || this.EMPTY_CELL;
      worksheet.getCell(`AF${row}`).value = item.altaTHP || this.EMPTY_CELL;
      worksheet.getCell(`AG${row}`).value =
        item.controlTorque || this.EMPTY_CELL;
      worksheet.getCell(`AH${row}`).value = item.altoTorque || this.EMPTY_CELL;
      worksheet.getCell(`AI${row}`).value = item.bajoTorque || this.EMPTY_CELL;
      worksheet.getCell(`AJ${row}`).value =
        item.corrienteMaxima || this.EMPTY_CELL;
      worksheet.getCell(`AK${row}`).value =
        item.temperaturaVSD || this.EMPTY_CELL;
      worksheet.getCell(`AL${row}`).value = item.manometrica || this.EMPTY_CELL;
      worksheet.getCell(`AM${row}`).value =
        item.antecedentes || this.EMPTY_CELL;
      worksheet.getCell(`AN${row}`).value =
        item.observaciones || this.EMPTY_CELL;
      row++;
    });

    const uuid = randomUUID();
    const path = `./public/excel/reporte-${uuid}.xlsx`;
    await workbook.xlsx.writeFile(path);
    return path;
  }
}

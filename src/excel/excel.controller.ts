import {
  BadRequestException,
  Body,
  Controller,
  Header,
  Post,
  StreamableFile,
} from '@nestjs/common';
import { ExcelService } from './excel.service';
import { createReadStream } from 'fs';

@Controller('excel')
export class ExcelController {
  constructor(private readonly excelService: ExcelService) {}

  @Post('create')
  @Header(
    'Content-Type',
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  )
  @Header('Content-Disposition', 'attachment; filename=reporte.xlsx')
  async createExcel(@Body() body: { data: Data }) {
    try {
      this.excelService.validateBody(body);
      const { data } = body;
      const path = await this.excelService.createExcel(data);
      const file = createReadStream(path);
      return new StreamableFile(file);
    } catch (error) {
      throw new BadRequestException(error.message);
    }
  }
}

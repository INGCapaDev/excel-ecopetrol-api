import { Controller, Post } from '@nestjs/common';
import { ExcelService } from './excel.service';

@Controller('excel')
export class ExcelController {
  constructor(private readonly excelService: ExcelService) {}

  @Post('create')
  async createExcel() {
    return 'create excel';
  }
}

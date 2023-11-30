import { BadRequestException, Body, Controller, Post } from '@nestjs/common';
import { ExcelService } from './excel.service';

@Controller('excel')
export class ExcelController {
  constructor(private readonly excelService: ExcelService) {}

  @Post('create')
  async createExcel(@Body() body: { data: Data }) {
    try {
      this.excelService.validateBody(body);
      const { data } = body;
      const path = this.excelService.createExcel(data);
      return path;
    } catch (error) {
      throw new BadRequestException(error.message);
    }
  }
}

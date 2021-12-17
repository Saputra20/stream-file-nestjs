import {
  Controller,
  Post,
  UseInterceptors,
  UploadedFile,
} from '@nestjs/common';
import { MiscService } from './misc.service';
import { FileInterceptor } from '@nestjs/platform-express';
import { ReadStream, createReadStream } from 'fs';
import * as xlsx from 'xlsx';
import { WorkBook, WorkSheet } from 'xlsx';
import { diskStorage } from 'multer';
import { fileName } from 'src/common/helper/multer';
import { exit } from 'process';

@Controller('misc')
export class MiscController {
  constructor(private readonly miscService: MiscService) {}

  @Post('upload')
  @UseInterceptors(
    FileInterceptor('file', {
      storage: diskStorage({
        destination: './files',
        filename: fileName,
      }),
    }),
  )
  async create(@UploadedFile() file: Express.Multer.File) {
    const wb: WorkBook = await new Promise((resolve, reject) => {
      const stream = createReadStream(file.path);
      const buffers = [];

      stream.on('data', (d) => buffers.push(d));

      stream.on('end', () => {
        const buffer = Buffer.concat(buffers);
        resolve(xlsx.read(buffer, { type: 'buffer' }));
      });

      stream.on('error', (error) => reject(error));
    });

    const sheetNames = wb.SheetNames;
    const sheet: WorkSheet = wb.Sheets[sheetNames[1]];
    const range = xlsx.utils.decode_range(sheet['!ref']);

    const students = [];
    for (let Row = range.s.r; Row <= range.e.r; ++Row) {
      if (Row === 0 || !sheet[xlsx.utils.encode_cell({ c: 0, r: Row })]) {
        continue;
      }
      const student = {
        pmbid: sheet[xlsx.utils.encode_cell({ c: 1, r: Row })].v,
        name: sheet[xlsx.utils.encode_cell({ c: 2, r: Row })].v,
        nim: sheet[xlsx.utils.encode_cell({ c: 3, r: Row })].v,
        status: sheet[xlsx.utils.encode_cell({ c: 4, r: Row })].v,
        paid: sheet[xlsx.utils.encode_cell({ c: 5, r: Row })].v,
        address: sheet[xlsx.utils.encode_cell({ c: 10, r: Row })].v,
      };

      students.push(student);
    }

    return students;
  }
}

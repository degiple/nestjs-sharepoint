import { Controller, Get, Query } from '@nestjs/common';
import { SharePointService } from './sharepoint.service';

@Controller('sharepoint')
export class SharePointController {
  constructor(private readonly sharePointService: SharePointService) {}

  @Get()
  async search(
    @Query('title') title: string,
    @Query('author') author: string,
    @Query('contenttype') contenttype: string,
    @Query('operator') operator = 'AND',
  ) {
    let query = '';
    if (title) query += `title:${title} `;
    if (author) query += `${operator} author:${author} `;
    if (contenttype) query += `${operator} contenttype:${contenttype}`;

    return await this.sharePointService.search(query.trim());
  }
}

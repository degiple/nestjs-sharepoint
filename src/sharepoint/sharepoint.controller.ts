// sharepoint.controller.ts
import { Controller, Get } from '@nestjs/common';
import { SharePointService } from './sharepoint.service';

@Controller('sharepoint')
export class SharePointController {
  constructor(private readonly sharePointService: SharePointService) {}

  @Get()
  async getAllSites(): Promise<any> {
    return this.sharePointService.getAllSites();
  }
}

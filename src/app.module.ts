import { Module } from '@nestjs/common';
import { AppController } from './app.controller';
import { AppService } from './app.service';
import { SharePointModule } from './sharepoint/sharepoint.module';

@Module({
  imports: [SharePointModule],
  controllers: [AppController],
  providers: [AppService],
})
export class AppModule {}

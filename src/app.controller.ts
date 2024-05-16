import { Controller, Get } from '@nestjs/common';
import { AppService } from './app.service';

@Controller()
export class AppController {
  constructor(private readonly appService: AppService) {}

  @Get('start-spider')
  startSpider() {
    this.appService.crawlAndSendEmail();
    return '爬虫已启动';
  }
}

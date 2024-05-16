import { Inject, Injectable } from '@nestjs/common';
import puppeteer from 'puppeteer';
import * as cron from 'node-cron';
import * as ExcelJS from 'exceljs';
import * as nodemailer from 'nodemailer';
import { EntityManager } from 'typeorm';

@Injectable()
export class AppService {
  private readonly mailTransporter = nodemailer.createTransport({
    // 配置您的电子邮件传输选项
    // 例如, 对于 SMTP:
    host: 'smtp.qq.com',
    port: 587,
    secure: false,
    auth: {
      user: '1047918517@qq.com',
      pass: 'dpbfhghslzpcbcai',
    },
  });

  async onModuleInit() {
    await this.crawlAndSendEmail();
  }

  //   constructor() {
  // 每天8点执行爬虫和发送邮件任务
  // cron.schedule('0 8 * * *', async () => {
  //   await this.crawlAndSendEmail();
  // });
  //     this.crawlAndSendEmail();
  //   }

  async crawlAndSendEmail() {
    const browser = await puppeteer.launch({
      headless: false,
      defaultViewport: {
        width: 0,
        height: 0,
      },
    });

    const page = await browser.newPage();

    await page.goto(
      'https://www.zhipin.com/web/geek/job?query=%E5%89%8D%E7%AB%AF&city=101280600&stage=807,808&salary=406',
    );

    await page.waitForSelector('.job-list-box');

    const totalPage = await page.$eval(
      '.options-pages a:nth-last-child(2)',
      (e) => {
        return parseInt(e.textContent);
      },
    );

    const allJobs = [];
    for (let i = 1; i <= 2; i++) {
      await page.goto(
        'https://www.zhipin.com/web/geek/job?query=%E5%89%8D%E7%AB%AF&city=101280600&stage=807,808&salary=406&page=' +
          i,
      );

      await page.waitForSelector('.job-list-box');

      const jobs = await page.$eval('.job-list-box', (el) => {
        return [...el.querySelectorAll('.job-card-wrapper')].map((item) => {
          return {
            job: {
              name: item.querySelector('.job-name').textContent,
              area: item.querySelector('.job-area').textContent,
              salary: item.querySelector('.salary').textContent,
            },
            link: item.querySelector('a').href,
            company: {
              name: item.querySelector('.company-name').textContent,
            },
          };
        });
      });
      allJobs.push(...jobs);
    }

    // console.log(allJobs);

    const allJobDetails = [];
    for (let i = 0; i < allJobs.length; i++) {
      await page.goto(allJobs[i].link);

      try {
        await page.waitForSelector('.job-sec-text');

        const jd = await page.$eval('.job-sec-text', (el) => {
          return el.textContent;
        });
        allJobs[i].desc = jd;

        // console.log(allJobs[i]);

        const job: any = {};

        job.name = allJobs[i].job.name;
        job.area = allJobs[i].job.area;
        job.salary = allJobs[i].job.salary;
        job.link = allJobs[i].link;
        job.company = allJobs[i].company.name;
        job.desc = allJobs[i].desc;

        allJobDetails.push(job);
      } catch (e) {}
    }

    // 创建 Excel 工作簿
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Recruitment Information');

    // 设置列宽
    worksheet.getColumn('A').width = 15; // 名称
    worksheet.getColumn('B').width = 30; // 区域
    worksheet.getColumn('C').width = 10; // 薪资
    worksheet.getColumn('D').width = 10; // 公司
    worksheet.getColumn('E').width = 60; // 描述
    worksheet.getColumn('F').width = 10; // 链接

    // 添加表头
    const headerRow = worksheet.addRow([
      '名称',
      '区域',
      '薪资',
      '公司',
      '描述',
      '链接',
    ]);

    headerRow.font = { bold: true };

    // 添加招聘信息数据
    allJobDetails.forEach((job) => {
      const row = worksheet.addRow([
        job.name,
        job.area,
        job.salary,
        job.company,
        job.desc,
        { text: '详情', hyperlink: job.link },
      ]);

      row.getCell('F').font = { color: { argb: '0000FF' }, underline: true }; // 设置链接字体为蓝色和下划线
    });

    // 将工作簿写入缓冲区
    const buffer = await workbook.xlsx.writeBuffer();

    // 发送电子邮件
    const mailOptions = {
      from: '1047918517@qq.com',
      to: 'jimmy_zhan2022@163.com',
      subject: 'Boss 直聘前端招聘信息',
      text: '请查看附件中的 Boss 直聘前端招聘信息。',
      attachments: [
        {
          filename: 'recruitment_information.xlsx',
          content: buffer,
        },
      ],
    };

    await this.mailTransporter.sendMail(mailOptions);

    await browser.close();
  }

  @Inject(EntityManager)
  private entityManager: EntityManager;
}

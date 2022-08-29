import {ls, mkdir} from 'shelljs';
import {join, parse} from 'path';
import {promises} from 'fs';
import {convert} from 'libreoffice-convert';
import {promisify} from 'util';

const convertAsync = promisify(convert);

const main = async () => {
  try {
    const dirs = ls('-RA', process.cwd());

    for (const entry of dirs) {
      if (entry.includes('.doc')) {
        const parsed = parse(entry);
        const oudDir = join('output', parsed.dir);

        mkdir('-p', oudDir);
        await docxToPdf(entry, join(oudDir, `${parsed.name}.pdf`));
      }
    }
  } catch (error) {
    console.error('Error Converting Files...');
    console.error(error);
  }
};

const docxToPdf = async (inputPath: string, outputPath: string) => {
  const ext = '.pdf';
  console.log(`Converting file ${inputPath} to PDF...`);

  const docxBuf = await promises.readFile(inputPath);

  let pdfBuf = await convertAsync(docxBuf, ext, undefined);

  console.log(`File converted. Saving File to ${outputPath}...`);
  await promises.writeFile(outputPath, pdfBuf);

  console.log('File Saved.');
};

main();

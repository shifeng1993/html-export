<template>
  <div class="top-bar">
    <button @click="downloadHtml('xx')">点击下载html</button>
    <button @click="downloadPDF('xx')">点击下载pdf</button>
    <button @click="downloadDoc('xx')">点击下载word</button>
  </div>
  <div ref="renderPageRef">
    <div class="page-container">
      <p v-for="i in 100" :key="i">
        这是预览的页面 我是里面的dom元素 这是预览的页面 我是里面的dom元素
      </p>
    </div>
  </div>
</template>

<script>
import { defineComponent, ref } from "vue";
// import axios from "axios";
import writer from "file-writer";
import { resumecss } from "./rendePage.css.js";

import html2Canvas from "html2canvas";
import JsPDF from "jspdf";

import Docxtemplater from "docxtemplater";
import PizZip from "pizzip";
import PizZipUtils from "pizzip/utils/index.js";
import { saveAs } from "file-saver";

function loadFile(url, callback) {
  PizZipUtils.getBinaryContent(url, callback);
}

export default defineComponent({
  name: "ExportPage",
  setup() {
    const renderPageRef = ref();
    const getHtmlContent = () => {
      const template = renderPageRef.value.innerHTML;
      let html = `<!DOCTYPE html>
                <html>
                <head>
                    <meta charset="utf-8">
                    <meta name="viewport" content="width=device-width,initial-scale=1.0">
                    <title>${"测试导出"}</title>
                    <style>
                      ${resumecss}
                    </style>
                </head>
                <body>
                    <div class="resume_preview_page" style="margin:0 auto;width:1200px">
                    ${template}
                    </div>
                </body>
                </html>`;
      return html;
    };

    const downloadHtml = (name) => {
      let html = getHtmlContent();
      writer(`${name}的简历.html`, html, "utf-8");
    };

    const downloadPDF = (name) => {
      html2Canvas(renderPageRef.value, {
        // allowTaint: true
        useCORS: true, //看情况选用上面还是下面的，
      }).then(function (canvas) {
        let contentWidth = canvas.width;
        let contentHeight = canvas.height;
        let pageHeight = (contentWidth / 592.28) * 841.89;
        let leftHeight = contentHeight;
        let position = 0;
        let imgWidth = 595.28;
        let imgHeight = (592.28 / contentWidth) * contentHeight;
        let pageData = canvas.toDataURL("image/jpeg", 1.0);
        let PDF = new JsPDF("", "pt", "a4");
        if (leftHeight < pageHeight) {
          PDF.addImage(pageData, "JPEG", 0, 0, imgWidth, imgHeight);
        } else {
          while (leftHeight > 0) {
            PDF.addImage(pageData, "JPEG", 0, position, imgWidth, imgHeight);
            leftHeight -= pageHeight;
            position -= 841.89;
            if (leftHeight > 0) {
              PDF.addPage();
            }
          }
        }
        PDF.save(name + ".pdf");
      });
    };

    const downloadDoc = () => {
      loadFile(
        `${window.location.origin}/html.docx`,
        function (error, content) {
          if (error) {
            throw error;
          }
          console.log(content);
          const zip = new PizZip(content);
          const doc = new Docxtemplater(zip, {
            paragraphLoop: true,
            linebreaks: true,
          });

          const template = renderPageRef.value.innerHTML;

          doc.setData({
            html: template,
          });

          // render the document (replace all occurences of {first_name} by John, {last_name} by Doe, ...)
          doc.render();

          const out = doc.getZip().generate({
            type: "blob",
            mimeType:
              "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
          });
          // Output the document using Data-URI
          saveAs(out, "output.docx");
        }
      );
    };
    return {
      renderPageRef,
      downloadHtml,
      downloadPDF,
      downloadDoc,
    };
  },
});
</script>

<style lang="less">
.page-container {
  margin: 0 auto;
  width: 794px;
  border: 1px solid #000;
  padding: 60px 40px;
}
</style>

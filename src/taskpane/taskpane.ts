/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// images references in the manifest
import "../../assets/icon-16.png";
import "../../assets/icon-32.png";
import "../../assets/icon-80.png";

/* global document, Office */

Office.onReady(info => {
  if (info.host === Office.HostType.OneNote) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

export async function run() {
  try {
    await OneNote.run((context) => {      
      const app = context.application

      const pages = app.getActiveSection().pages;

      pages.load('id,title');

      return context.sync()
        .then(async () => {
          return await pages.items.reduce(
            async (promiseChain, page) =>
              await promiseChain.then(
                async () => {
                  var pageId = page.id;
                  var pageTitle = page.title;
                  console.log(pageTitle + ': ' + pageId);

                  return await processPage(page)
                }

              ),
            Promise.resolve()
          )
        })
        .catch((error) => {
          // app.showNotification("Error: " + error);
          console.log("Error: " + error);
          if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
          }
        });
    });
  } catch (error) {
    console.log("Error: " + error);
  }
}

async function processPage(page: OneNote.Page): Promise<any> {
  const context = page.context

  const app = context.application
  app.navigateToPage(page)
  console.log('processPage', page.title)

  return context.sync()
    .then(() => {
      const activePage = app.getActivePage()
      activePage.track()
      activePage.load('id,title')

      return context.sync()
        .then(async () => {
          await exportPage(activePage)
        })
        .finally(() => {
          activePage.untrack()
        })
    })
}

async function exportPage(page: OneNote.Page): Promise<any> {
  const context = page.context

  const pageContents = page.contents;
  pageContents.load('id,type');
  console.log('exportPage', page.title)

  return context.sync()
    .then(async () => {
      console.log('pageContents', page.title, pageContents.count, pageContents.items)

      return await pageContents.items.reduce(
        async (promiseChain, content) =>
          promiseChain.then(
            async () => {
              await processContent(content)
            }
          ),
        Promise.resolve()
      )
    })
}

async function processContent(content: OneNote.PageContent): Promise<any> {
  const context = content.context

  console.log('processContent', content.type)

  switch (content.type) {
    case 'Outline':
      return await processOutline(content.outline)
  }
}

async function processOutline(outline: OneNote.Outline): Promise<any> {
  const context = outline.context

  console.log('processOutline', outline)
  outline.track()
  outline.load('id,type,paragraphs')

  return context.sync()
    .then(async () => {
      return await outline.paragraphs.items.reduce(
        async (promiseChain, paragraph) => {
          paragraph.track()
          return promiseChain
            .then(
              async () => {
                return await processParagraph(paragraph)
              }
            )
            .finally(() => {
              paragraph.untrack()
            })
        },
        Promise.resolve()
      )
    })
    .finally(() => {
      outline.untrack()
    })
}

async function processParagraph(paragraph: OneNote.Paragraph): Promise<any> {
  const context = paragraph.context

  console.log('processParagraph', paragraph.type)

  switch (paragraph.type) {
    case 'RichText': {
      paragraph.load('richtext')
      return context.sync()
        .then(async () => {
          return await processRichText(paragraph.richText)
        })
    }
    case 'Image':
      paragraph.load('image')
      return context.sync()
        .then(async () => {
          return await processImage(paragraph.image)
        })

    case 'Ink':
      paragraph.load('inkwords')
      return context.sync()
        .then(async () => {
          return await processInk(paragraph.inkWords)
        })

    case 'Table':
      return await processTable(paragraph.table)

    case 'Other':
      return Promise.resolve(false)
  }
}

async function processRichText(richText: OneNote.RichText): Promise<any> {
  const context = richText.context

  console.log('processRichText', richText)

  richText.track()
  richText.load('text')
  const html = richText.getHtml()

  return context.sync()
    .then(async () => {
      console.log('richText.text', richText.text)
      console.log('richText.html', html.value)
    })
    .finally(() => {
      richText.untrack()
    })

}

async function processImage(image: OneNote.Image): Promise<any> {
  const context = image.context

  console.log('processImage', image)
  image.track()
  image.load('link,hyperlink,description,width,height')
  const base64Image = image.getBase64Image()

  return context.sync()
    .then(() => {
      console.log('image.hyperlink', image.hyperlink)
      console.log('image.description', image.description)
      console.log('image.width', image.width)
      console.log('image.height', image.height)
      // console.log('image.base64', base64Image.value)
    })
    .finally(() => {
      image.untrack()
    })
}

async function processInk(ink: OneNote.InkWordCollection): Promise<any> {
  console.log('processInk', ink)
}

async function processTable(table: OneNote.Table): Promise<any> {
  console.log('processTable', table)
}

name: Remove Event's handle
description: Test remove Excel Events
author: Shanbo
host: EXCEL
api_set: {}

script:

    content: |
          $("#add").click(() => tryCatch(add));
          $("#remove").click(() => tryCatch(remove));

          var eventResult;
          async function add() {
              await Excel.run(async (ctx) => {
              var sheetName = "Sheet1";
              var worksheet = ctx.workbook.worksheets.getItem(sheetName);
              eventResult = worksheet.onSelectionChanged.add(onSelectionChanged);
              await ctx.sync();
              OfficeHelpers.UI.notify("add event sucess");
              });

            function onSelectionChanged(args) {
              console.log("selection changed " + args.address);
            }

          }


          async function remove() {
            if(!eventResult || !eventResult.context)
            {
              console.log("no added the handler");
              return;
            }

            await Excel.run(eventResult.context, async (ctx) => {
                eventResult.remove();
                await ctx.sync() ;
                OfficeHelpers.UI.notify("remove event sucess");
            });
          }


          /** Default helper for invoking an action and handling errors. */
          async function tryCatch(callback) {
              try {
                  await callback();
              }
              catch (error) {
                  OfficeHelpers.UI.notify(error);
                  OfficeHelpers.Utilities.log(error);
              }
          }

    language: typescript

template:

    content: |
      <button id="add" class="ms-Button">
          <span class="ms-Button-label">add</span>
      </button>

      <button id="remove" class="ms-Button">
          <span class="ms-Button-label">remove</span>
      </button>

    language: html

style:

    content: "/* Your style goes here */\r\n"

    language: css

libraries: |

    https://appsforoffice.microsoft.com/lib/beta/hosted/office.debug.js

    https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts



    office-ui-fabric-js@1.4.0/dist/css/fabric.min.css

    office-ui-fabric-js@1.4.0/dist/css/fabric.components.min.css



    core-js@2.4.1/client/core.min.js

    @types/core-js



    @microsoft/office-js-helpers@0.7.4/dist/office.helpers.min.js

    @microsoft/office-js-helpers@0.7.4/dist/office.helpers.d.ts



    jquery@3.1.1

    @types/jquery

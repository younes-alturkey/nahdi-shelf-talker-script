/*
 * Authors
 * Mohammed Alsayed
 * Younes Alturkey
 * Sarah Alshawkani
 */

Main();

function Main() {
  function start(path) {
    // check if there is a file
    if (path != null) {
      // render on existing open document
      var doc = app.documents.item(0);

      // to unify the script measerment in all devices
      doc.viewPreferences.horizontalMeasurementUnits =
        MeasurementUnits.MILLIMETERS;
      doc.viewPreferences.verticalMeasurementUnits =
        MeasurementUnits.MILLIMETERS;

      // This should be big enouch to contain the size of the wide and normal templates
      doc.documentPreferences.pageWidth = '310 mm';
      doc.documentPreferences.pageHeight = '397 mm';

      // to identfy the color for descrption for each group
      var color = doc.colors.add({
        name: 'C=0 M=0 Y=0 K=0',
        space: ColorSpace.CMYK,
        model: ColorModel.process,
        colorValue: [94, 58, 52, 37],
      });

      // adding white color to doc
      doc.colors.add({
        name: 'white',
        model: ColorModel.process,
        colorValue: [0, 0, 0, 0],
      });

      // Each textframe contains a whole Excel column
      var tmp_textframe = doc.pages[0].textFrames.add();
      tmp_textframe.properties = {
        textFramePreferences: {
          autoSizingReferencePoint: 1953459301,
          autoSizingType: 1752070000,
        },
      };

      var tmp_textframe1 = doc.pages[0].textFrames.add();
      tmp_textframe1.properties = {
        textFramePreferences: {
          autoSizingReferencePoint: 1953459301,
          autoSizingType: 1752070000,
        },
      };

      var tmp_textframe2 = doc.pages[0].textFrames.add();
      tmp_textframe2.properties = {
        textFramePreferences: {
          autoSizingReferencePoint: 1953459301,
          autoSizingType: 1752070000,
        },
      };

      var tmp_textframe3 = doc.pages[0].textFrames.add();
      tmp_textframe3.properties = {
        textFramePreferences: {
          autoSizingReferencePoint: 1953459301,
          autoSizingType: 1752070000,
        },
      };

      var tmp_textframe4 = doc.pages[0].textFrames.add();
      tmp_textframe4.properties = {
        textFramePreferences: {
          autoSizingReferencePoint: 1953459301,
          autoSizingType: 1752070000,
        },
      };

      var tmp_textframe5 = doc.pages[0].textFrames.add();
      tmp_textframe5.properties = {
        textFramePreferences: {
          autoSizingReferencePoint: 1953459301,
          autoSizingType: 1752070000,
        },
      };

      var tmp_textframe6 = doc.pages[0].textFrames.add();
      tmp_textframe6.properties = {
        textFramePreferences: {
          autoSizingReferencePoint: 1953459301,
          autoSizingType: 1752070000,
        },
      };
      var tmp_textframe7 = doc.pages[0].textFrames.add();
      tmp_textframe7.properties = {
        textFramePreferences: {
          autoSizingReferencePoint: 1953459301,
          autoSizingType: 1752070000,
        },
      };

      var tmp_textframe8 = doc.pages[0].textFrames.add();
      tmp_textframe8.properties = {
        textFramePreferences: {
          autoSizingReferencePoint: 1953459301,
          autoSizingType: 1752070000,
        },
      };

      var tmp_textframe9 = doc.pages[0].textFrames.add();
      tmp_textframe9.properties = {
        textFramePreferences: {
          autoSizingReferencePoint: 1953459301,
          autoSizingType: 1752070000,
        },
      };

      var tmp_textframe10 = doc.pages[0].textFrames.add();
      tmp_textframe10.properties = {
        textFramePreferences: {
          autoSizingReferencePoint: 1953459301,
          autoSizingType: 1752070000,
        },
      };

      //Automate Excel Sheet Cell Range Selection
      function setExcelImportPrefs(maxRange, letter) {
        //A1:A30 for exmaple if the maxRange is 30
        app.excelImportPreferences.rangeName = letter
          .concat('1:')
          .concat(letter)
          .concat(maxRange);
        app.excelImportPreferences.tableFormatting =
          TableFormattingOptions.excelUnformattedTabbedText;
      }

      try {
        // Take number of rows from user to automate cell range input A1:A[maxRange]
        maxRange = maxRange = prompt(
          'Please, enter the max number of rows (e.g. 30).',
          '',
          'Excel Sheet Max Range For Rows'
        );

        // read NAME column
        tmp_textframe.place(path, setExcelImportPrefs(maxRange, 'A'));

        // read ITEM column
        tmp_textframe1.place(path, setExcelImportPrefs(maxRange, 'B'));

        //read Offer Description column
        tmp_textframe2.place(path, setExcelImportPrefs(maxRange, 'C'));

        // read COMPONENT_DESC column
        tmp_textframe3.place(path, setExcelImportPrefs(maxRange, 'D'));

        // read Group retail Price column
        tmp_textframe4.place(path, setExcelImportPrefs(maxRange, 'E'));

        //  read Saving column
        tmp_textframe5.place(path, setExcelImportPrefs(maxRange, 'F'));

        // read Saving column
        tmp_textframe6.place(path, setExcelImportPrefs(maxRange, 'G'));

        // read % Discount column
        tmp_textframe7.place(path, setExcelImportPrefs(maxRange, 'H'));

        // read Date column
        tmp_textframe8.place(path, setExcelImportPrefs(maxRange, 'I'));

        // read Currency column
        tmp_textframe9.place(path, setExcelImportPrefs(maxRange, 'J'));

        // read Temple column
        tmp_textframe10.place(path, setExcelImportPrefs(maxRange, 'K'));

        // pass the data to array, each array cell contains an excel line
        var namesArray = tmp_textframe.parentStory.contents.split('\r');
        var itemsArray = tmp_textframe1.parentStory.contents.split('\r');
        var offerDescriptionArray = tmp_textframe2.parentStory.contents.split(
          '\r'
        );
        var componentDescriptionArray = tmp_textframe3.parentStory.contents.split(
          '\r'
        );
        var retailPriceArray = tmp_textframe4.parentStory.contents.split('\r');
        var promoPriceArray = tmp_textframe5.parentStory.contents.split('\r');
        var savingAmountArray = tmp_textframe6.parentStory.contents.split('\r');
        var discountsArray = tmp_textframe7.parentStory.contents.split('\r');
        var datesArray = tmp_textframe8.parentStory.contents.split('\r');
        var currencyArray = tmp_textframe9.parentStory.contents.split('\r');
        var templateTypeArray = tmp_textframe10.parentStory.contents.split(
          '\r'
        );

        //Remove frames after data extraction
        tmp_textframe.remove();
        tmp_textframe1.remove();
        tmp_textframe2.remove();
        tmp_textframe3.remove();
        tmp_textframe4.remove();
        tmp_textframe5.remove();
        tmp_textframe6.remove();
        tmp_textframe7.remove();
        tmp_textframe8.remove();
        tmp_textframe9.remove();
        tmp_textframe10.remove();

        //read from the base graphics folder should have the two templates and required images like slash and X2
        var baseGraphicsFolder = Folder.selectDialog(
          'Please, select the base images folder.'
        );

        // Get all the files in the selected folder
        var baseGraphicsFiles = baseGraphicsFolder.getFiles();

        //read from the products images
        var productsImagesFolder = Folder.selectDialog(
          'Please, select the products images folder.'
        );

        // Get all the files in the selected folder

        var productsImages = productsImagesFolder.getFiles();

        function getImageFileFromSKU(sku) {
          for (var i = 0; i < productsImages.length; i++) {
            if (productsImages[i].toString().indexOf(sku.toString()) >= 0) {
              return productsImages[i];
            }
          }

          return productsImages[0];
        }

        // Get .ai base template images to render on document
        // The folder contains 0-wide-template.ai, 1-normal-template.ai, 2-slash.ai, 3-x2.ai ORDER is important
        var shelfTalkerWideBase = baseGraphicsFiles[0];
        var shelfTalkerNormalBase = baseGraphicsFiles[1];
        var shelfTalkerRetailPriceSlash = baseGraphicsFiles[2];
        var shelfTalkerX2 = baseGraphicsFiles[3];

        //[y, x, y+h, x+w]
        //Default positions for Wide template
        var widePosition = [38, 41, 112, 268];
        var wideImagePosition = [47.609, 193.77, 95.25, 230.588];
        var wideTextPosition = [95.25, 192.179, 106.25, 235.269];
        var wideDatePosition = [100.25, 242.25, 106.25, 266];

        // CASE: BUY x GET x FREE
        var wideRetailPricePosition = [75, 79.6, 94.903, 137];
        var widePromoPosition = [56.1, 89, 67.1, 123.6];
        var wideRetailPriceSlashPosition = [54.7, 88.4, 67.087, 123];
        var wideCurrencyPosition = [64, 120.333, 75, 137.333];

        // CASE: خصم % على الحبة الثانية
        var widePercentagePromoPosition = [58.151, 61, 91.849, 149.5];

        // CASE: 2 @
        var widePercentagePricePosition = [61.478, 67.1, 81.381, 123.6];
        var widePercentageX2Position = [82, 196, 91.405, 205.864];
        var widePercentagePriceCurrencyPosition = [63.704, 127, 79.155, 150.75];

        //Default positions for normal template
        var normalPosition = [38, 85.954, 144.616, 235.693];
        var normalImagePosition = [58.644, 172.859, 100.808, 226.707];
        var normalTextPosition = [112.016, 194.703, 126.404, 232.041];
        var normalDatePosition = [132.667, 88.333, 141.471, 135.635];
        var normalDirectDiscountPosition = [62.059, 94.333, 97.392, 153.338];

        // start at page 0
        var currentNumberOfRenderedProducts = 0;
        var pageIndex = 0;

        // Decides what kind of render for each products for the switch cases
        var offerType;

        //Render all products from the the Excel sheet
        for (var i = 1; i < maxRange; i++) {
          //Check which template type to render
          if (templateTypeArray[i].indexOf('Wide') >= 0) {
            // Render general info and images
            // This renders for all products regardless of offer type
            var wideBase = doc.pages[pageIndex].textFrames.add({
              geometricBounds: widePosition,
            });
            wideBase.place(shelfTalkerWideBase);
            wideBase.fit(FitOptions.FRAME_TO_CONTENT);
            wideBase.fit(FitOptions.PROPORTIONALLY);

            var wideImage = doc.pages[pageIndex].textFrames.add({
              geometricBounds: wideImagePosition,
            });
            wideImage.place(getImageFileFromSKU(itemsArray[i]));
            wideImage.fit(FitOptions.PROPORTIONALLY);

            var wideDescription = doc.pages[pageIndex].textFrames.add({
              geometricBounds: wideTextPosition,
            });
            //changing the type of font for the wideDescription and resize it
            wideDescription.texts[0].appliedFont = 'Nahdi	Black';
            wideDescription.texts[0].pointSize = 6;
            wideDescription.texts[0].parentStory.justification =
              Justification.CENTER_ALIGN;
            wideDescription.texts[0].fillColor = color;
            wideDescription.contents = offerDescriptionArray[i].toString();

            //changing the type of font for the wideDate and resize it
            // You need to set a font that displays the content in English Nahdi Black is Arabic based and does't render correctly
            // try Nahdi Black to see the issue
            var wideDate = doc.pages[pageIndex].textFrames.add({
              geometricBounds: wideDatePosition,
            });
            wideDate.texts[0].appliedFont = app.fonts.item('Segoe UI Emoji');
            wideDate.texts[0].appliedLanguage = 'English: USA';
            wideDate.texts[0].pointSize = 9;
            wideDate.texts[0].parentStory.justification =
              Justification.CENTER_ALIGN;
            wideDate.texts[0].fillColor = doc.colors.item('white');
            wideDate.contents = datesArray[i].toString();

            //Render description specific info and images

            //Check promo type to render according elements
            if (componentDescriptionArray[i].toString().indexOf('FREE') >= 0)
              offerType = 1;
            else if (componentDescriptionArray[i].toString().indexOf('%') >= 0)
              offerType = 2;
            else if (componentDescriptionArray[i].toString().indexOf('@') >= 0)
              offerType = 3;
            else if (
              componentDescriptionArray[i].toString().indexOf('B/G') >= 0
            )
              offerType = 4;
            else offerType = -1;

            switch (offerType) {
              // CASE: BUY x GET x FREE
              case 1:
                var wideRetailPrice = doc.pages[pageIndex].textFrames.add({
                  geometricBounds: wideRetailPricePosition,
                });
                wideRetailPrice.texts[0].appliedFont = 'Nahdi	Black';
                wideRetailPrice.texts[0].pointSize = 60;
                wideRetailPrice.texts[0].parentStory.justification =
                  Justification.CENTER_ALIGN;
                wideRetailPrice.texts[0].fillColor = doc.colors.item('white');
                wideRetailPrice.contents = promoPriceArray[i];

                var widePromoPrice = doc.pages[pageIndex].textFrames.add({
                  geometricBounds: widePromoPosition,
                });
                widePromoPrice.texts[0].appliedFont = 'Nahdi	Black';
                widePromoPrice.texts[0].pointSize = 36;
                widePromoPrice.texts[0].parentStory.justification =
                  Justification.CENTER_ALIGN;
                widePromoPrice.texts[0].fillColor = doc.colors.item('white');
                widePromoPrice.contents = retailPriceArray[i];

                var wideCurrency = doc.pages[pageIndex].textFrames.add({
                  geometricBounds: wideCurrencyPosition,
                });
                wideCurrency.texts[0].appliedFont = 'Nahdi	Black';
                wideCurrency.texts[0].pointSize = 24;
                wideCurrency.texts[0].parentStory.justification =
                  Justification.CENTER_ALIGN;
                wideCurrency.texts[0].fillColor = doc.colors.item('white');
                wideCurrency.contents = currencyArray[i];

                var wideRetailPriceSlash = doc.pages[pageIndex].textFrames.add({
                  geometricBounds: wideRetailPriceSlashPosition,
                });
                wideRetailPriceSlash.place(shelfTalkerRetailPriceSlash);
                wideRetailPriceSlash.fit(FitOptions.PROPORTIONALLY);
                break;

              // CASE: خصم % على الحبة الثانية
              case 2:
                var widePercentagePromo = doc.pages[pageIndex].textFrames.add({
                  geometricBounds: widePercentagePromoPosition,
                });
                widePercentagePromo.texts[0].appliedFont = 'Nahdi	Black';
                widePercentagePromo.texts[0].pointSize = 38;
                widePercentagePromo.texts[0].parentStory.justification =
                  Justification.RIGHT_ALIGN;
                widePercentagePromo.texts[0].fillColor = doc.colors.item(
                  'white'
                );
                widePercentagePromo.contents = componentDescriptionArray[
                  i
                ].toString();
                break;

              // CASE: 2 @
              case 3:
                var widePercentagePrice = doc.pages[pageIndex].textFrames.add({
                  geometricBounds: widePercentagePricePosition,
                });
                widePercentagePrice.texts[0].appliedFont = 'Nahdi	Black';
                widePercentagePrice.texts[0].pointSize = 64;
                widePercentagePrice.texts[0].parentStory.justification =
                  Justification.RIGHT_ALIGN;
                widePercentagePrice.texts[0].fillColor = doc.colors.item(
                  'white'
                );
                widePercentagePrice.contents = promoPriceArray[i].toString();

                var widePercentagePriceCurrency = doc.pages[
                  pageIndex
                ].textFrames.add({
                  geometricBounds: widePercentagePriceCurrencyPosition,
                });
                widePercentagePriceCurrency.texts[0].appliedFont = 'Nahdi	Black';
                widePercentagePriceCurrency.texts[0].pointSize = 32;
                widePercentagePriceCurrency.texts[0].parentStory.justification =
                  Justification.CENTER_ALIGN;
                widePercentagePriceCurrency.texts[0].fillColor = doc.colors.item(
                  'white'
                );
                widePercentagePriceCurrency.contents = currencyArray[i];

                var widePercentageX2 = doc.pages[pageIndex].textFrames.add({
                  geometricBounds: widePercentageX2Position,
                });
                widePercentageX2.place(shelfTalkerX2);
                widePercentageX2.fit(FitOptions.PROPORTIONALLY);
                break;

              // CASE: B/G
              case 4:
                // ADD B/G OFFER ADDITIONAL RENDERING LOGIC HERE
                break;
              default:
                break;
            }
          } else {
            // If template is not set to Wide then I will render Normal template
            var normalBase = doc.pages[pageIndex].textFrames.add({
              geometricBounds: normalPosition,
            });
            normalBase.place(shelfTalkerNormalBase);
            normalBase.fit(FitOptions.FRAME_TO_CONTENT);
            normalBase.fit(FitOptions.PROPORTIONALLY);

            var normalImage = doc.pages[pageIndex].textFrames.add({
              geometricBounds: normalImagePosition,
            });
            normalImage.place(getImageFileFromSKU(itemsArray[i]));
            normalImage.fit(FitOptions.PROPORTIONALLY);

            var normalDescription = doc.pages[pageIndex].textFrames.add({
              geometricBounds: normalTextPosition,
            });
            normalDescription.texts[0].appliedFont = 'Nahdi	Black';
            normalDescription.texts[0].pointSize = 8;
            normalDescription.texts[0].parentStory.justification =
              Justification.CENTER_ALIGN;
            normalDescription.texts[0].fillColor = color;
            normalDescription.contents = offerDescriptionArray[i].toString();

            // You need to set a font that displays the content in English Nahdi Black is Arabic based and does't render correctly
            // try 'Nahdi Black' to see the issue
            var normalDate = doc.pages[pageIndex].textFrames.add({
              geometricBounds: normalDatePosition,
            });
            normalDate.texts[0].appliedFont = app.fonts.item('Segoe UI Emoji');
            normalDate.texts[0].appliedLanguage = 'English: USA';
            normalDate.texts[0].pointSize = 16;
            normalDate.texts[0].parentStory.justification =
              Justification.CENTER_ALIGN;
            normalDate.texts[0].fillColor = doc.colors.item('white');
            normalDate.contents = datesArray[i].toString();

            //Render description specific info and images
            if (
              componentDescriptionArray[i]
                .toString()
                .indexOf('Direct_Discount') >= 0
            )
              offerType = 1;
            else offerType = -1;

            switch (offerType) {
              // CASE: Direct_Discount
              case 1:
                var normalDirectDiscount = doc.pages[pageIndex].textFrames.add({
                  geometricBounds: normalDirectDiscountPosition,
                });
                normalDirectDiscount.texts[0].appliedFont = 'Nahdi	Black';
                normalDirectDiscount.texts[0].pointSize = 48;
                normalDirectDiscount.texts[0].parentStory.justification =
                  Justification.CENTER_ALIGN;
                normalDirectDiscount.texts[0].fillColor = doc.colors.item(
                  'white'
                );
                normalDirectDiscount.contents = savingAmountArray[i].toString();
                break;

              // You may set aditional cases for the normal template here case 2

              default:
                break;
            }
          }

          //Increment wide epositions
          widePosition = [
            widePosition[0] + 120,
            widePosition[1],
            widePosition[2] + 120,
            widePosition[3],
          ];

          wideImagePosition = [
            wideImagePosition[0] + 120,
            wideImagePosition[1],
            wideImagePosition[2] + 120,
            wideImagePosition[3],
          ];

          wideTextPosition = [
            wideTextPosition[0] + 120,
            wideTextPosition[1],
            wideTextPosition[2] + 120,
            wideTextPosition[3],
          ];

          var wideDatePosition = [
            wideDatePosition[0] + 120,
            wideDatePosition[1],
            wideDatePosition[2] + 120,
            wideDatePosition[3],
          ];

          var wideRetailPricePosition = [
            wideRetailPricePosition[0] + 120,
            wideRetailPricePosition[1],
            wideRetailPricePosition[2] + 120,
            wideRetailPricePosition[3],
          ];

          var widePromoPosition = [
            widePromoPosition[0] + 120,
            widePromoPosition[1],
            widePromoPosition[2] + 120,
            widePromoPosition[3],
          ];

          var wideRetailPriceSlashPosition = [
            wideRetailPriceSlashPosition[0] + 120,
            wideRetailPriceSlashPosition[1],
            wideRetailPriceSlashPosition[2] + 120,
            wideRetailPriceSlashPosition[3],
          ];

          var wideCurrencyPosition = [
            wideCurrencyPosition[0] + 120,
            wideCurrencyPosition[1],
            wideCurrencyPosition[2] + 120,
            wideCurrencyPosition[3],
          ];

          var widePercentagePromoPosition = [
            widePercentagePromoPosition[0] + 120,
            widePercentagePromoPosition[1],
            widePercentagePromoPosition[2] + 120,
            widePercentagePromoPosition[3],
          ];

          var widePercentagePricePosition = [
            widePercentagePricePosition[0] + 120,
            widePercentagePricePosition[1],
            widePercentagePricePosition[2] + 120,
            widePercentagePricePosition[3],
          ];

          var widePercentageX2Position = [
            widePercentageX2Position[0] + 120,
            widePercentageX2Position[1],
            widePercentageX2Position[2] + 120,
            widePercentageX2Position[3],
          ];

          var widePercentagePriceCurrencyPosition = [
            widePercentagePriceCurrencyPosition[0] + 120,
            widePercentagePriceCurrencyPosition[1],
            widePercentagePriceCurrencyPosition[2] + 120,
            widePercentagePriceCurrencyPosition[3],
          ];

          //Increment normal epositions
          normalPosition = [
            normalPosition[0] + 120,
            normalPosition[1],
            normalPosition[2] + 120,
            normalPosition[3],
          ];

          normalImagePosition = [
            normalImagePosition[0] + 120,
            normalImagePosition[1],
            normalImagePosition[2] + 120,
            normalImagePosition[3],
          ];

          normalTextPosition = [
            normalTextPosition[0] + 120,
            normalTextPosition[1],
            normalTextPosition[2] + 120,
            normalTextPosition[3],
          ];

          normalDatePosition = [
            normalDatePosition[0] + 120,
            normalDatePosition[1],
            normalDatePosition[2] + 120,
            normalDatePosition[3],
          ];

          normalDirectDiscountPosition = [
            normalDirectDiscountPosition[0] + 120,
            normalDirectDiscountPosition[1],
            normalDirectDiscountPosition[2] + 120,
            normalDirectDiscountPosition[3],
          ];

          currentNumberOfRenderedProducts++;

          // Only 3 products are allowed per page
          if (currentNumberOfRenderedProducts >= 3) {
            // create new page and increment page index
            doc.pages.add();
            pageIndex++;

            // reset wide positions after new page creation
            currentNumberOfRenderedProducts = 0;
            widePosition = [38, 41, 112, 268];
            wideImagePosition = [47.609, 193.77, 95.25, 230.588];
            wideTextPosition = [95.25, 192.179, 106.25, 235.269];
            wideDatePosition = [100.25, 242.25, 106.25, 266];
            wideRetailPricePosition = [75, 79.6, 94.903, 137];
            widePromoPosition = [56.1, 89, 67.1, 123.6];
            wideRetailPriceSlashPosition = [54.7, 88.4, 67.087, 123];
            wideCurrencyPosition = [64, 120.333, 75, 137.333];
            widePercentagePromoPosition = [58.151, 61, 91.849, 149.5];
            widePercentagePricePosition = [61.478, 67.1, 81.381, 123.6];
            widePercentageX2Position = [82, 196, 91.405, 205.864];
            widePercentagePriceCurrencyPosition = [63.704, 127, 79.155, 150.75];

            // reset normal positions after new page creation
            normalPosition = [38, 85.954, 144.616, 235.693];
            normalImagePosition = [58.644, 172.859, 100.808, 226.707];
            normalTextPosition = [112.016, 194.703, 126.404, 232.041];
            normalDatePosition = [132.667, 88.333, 141.471, 135.635];
            normalDirectDiscountPosition = [62.059, 94.333, 97.392, 153.338];
          }
        }
      } catch (e) {
        // Alert the user if an error occured
        alert('An error occured');
      }
    }
  }

  // calling the function && let the user set the desired excel file
  app.doScript(
    "start(File.openDialog('select file', '*.*'))",
    ScriptLanguage.JAVASCRIPT,
    undefined,
    UndoModes.ENTIRE_SCRIPT,
    'test'
  );
}

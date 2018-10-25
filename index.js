const reviewsCrawler = require("amazon-reviews-crawler");
const LineByLineReader = require("line-by-line");
const dateTime = require("node-datetime");
const dt = dateTime.create();
const formatted = dt.format("Y-m-d_H:M:S");
var ne = require("node-each");
var excel = require("excel4node");
var amazon = require("amazon-product-api");
var client = amazon.createClient({
  awsId: "XXXXXXXXXXXXXXXXXXXX",
  awsSecret: "XXXXXXXXXXXXXXXXXXXX"
});
lr = new LineByLineReader("amazonSource.txt"); //Script Load Product Ids From This File, Each Line Contain Single Product Id
var workbook = new excel.Workbook();
var worksheet = workbook.addWorksheet("Sheet 1");
const apid = [];
worksheet.cell(1, 1).string("ASIN");
worksheet.cell(1, 2).string("Model");
worksheet.cell(1, 3).string("Title");
worksheet.cell(1, 4).string("DetailPageURL");
worksheet.cell(1, 5).string("ListPrice");
worksheet.cell(1, 6).string("Price");
worksheet.cell(1, 7).string("AmountSaved");
worksheet.cell(1, 8).string("PercentageSaved");
worksheet.cell(1, 9).string("SalesRank");
worksheet.cell(1, 10).string("ProductGroup");
worksheet.cell(1, 11).string("ProductTypeName");
worksheet.cell(1, 12).string("UPC");
worksheet.cell(1, 13).string("Availability");
worksheet.cell(1, 14).string("IsEligibleForPrime");
worksheet.cell(1, 15).string("Manufacturer");
worksheet.cell(1, 16).string("Brand");
worksheet.cell(1, 17).string("Average Ratting");
worksheet.cell(1, 18).string("Total Rattings");
worksheet.cell(1, 19).string("Any Error In Processing");
var retryValue = 0;
var options = {
  debug: true,
  on: "time",
  when: 150.000001
};
lr.on("line", line => {
  const res = line.slice(line.indexOf("-") + 1);
  apid.push(res);
  console.log("Loading Product Ids");
});
lr.on("error", err => {
  console.log("Error: " + err);
});
lr.on("end", () => {
  crawler(apid);
});
crawler = toCrawl => {
  ne.each(
    toCrawl,
    (el, i) => {
      return new Promise(resolve => {
        setTimeout(() => {
          if (retryValue != 0) {
            el = retryValue;
          }
          reviewsCrawler(el)
            .then(results => {
              retryValue = 0;
              var totalRattingsByUsers = 0;
              var totalRattings = 0;
              var AverageRatting = 0;
              if (
                typeof results["reviews"] !== undefined &&
                results["reviews"] !== null
              ) {
                ne.each(results.reviews, (el, i) => {
                  if (el !== undefined) {
                    totalRattingsByUsers = totalRattingsByUsers + el.rating;
                    totalRattings = totalRattings + 1;
                  }
                }).then(() => {
                  if (totalRattings != 0) {
                    AverageRatting = totalRattingsByUsers / totalRattings;
                    AverageRatting = Math.round(AverageRatting * 100) / 100;
                  }
                  console.log(
                    "***************************************SCANNING NEW PRODUCT***************************************"
                  );
                  console.log(
                    "Number: " +
                      i +
                      " Product Id: " +
                      el +
                      " AverageRatting: " +
                      AverageRatting +
                      " Total Rattings: " +
                      totalRattings
                  );
                  worksheet.cell(i + 2, 1).string(el);
                  worksheet.cell(i + 2, 17).number(AverageRatting);
                  worksheet.cell(i + 2, 18).number(totalRattings);
                  client.itemLookup(
                    {
                      itemId: el,
                      responseGroup: "Large"
                    },
                    function(err, results, response) {
                      if (err) {
                        console.log(JSON.stringify(err));
                        worksheet.cell(i + 2, 19).string(JSON.stringify(err));
                        resolve();
                      } else {
                        try {
                          console.log(JSON.stringify(results[0]));
                          var ItemIntributes = null;
                          console.log("=>> ASIN: " + el);
                          if (results[0].ItemAttributes !== undefined) {
                            ItemIntributes = results[0].ItemAttributes[0];
                          }
                          if (ItemIntributes.Model !== undefined) {
                            console.log(
                              "=>> Model: " + ItemIntributes.Model[0]
                            );
                            worksheet
                              .cell(i + 2, 2)
                              .string(ItemIntributes.Model[0]);
                          }
                          if (ItemIntributes.Title !== undefined) {
                            console.log(
                              "=>> Title: " + ItemIntributes.Title[0]
                            );
                            worksheet
                              .cell(i + 2, 3)
                              .string(ItemIntributes.Title[0]);
                          }
                          if (results[0].DetailPageURL !== undefined) {
                            console.log(
                              "=>> DetailPageURL: " +
                                results[0].DetailPageURL[0]
                            );
                            worksheet
                              .cell(i + 2, 4)
                              .string(results[0].DetailPageURL[0]);
                          }
                          if (ItemIntributes.ListPrice !== undefined) {
                            console.log(
                              "=>> ListPrice: " +
                                ItemIntributes.ListPrice[0].FormattedPrice[0]
                            );
                            worksheet
                              .cell(i + 2, 5)
                              .string(
                                ItemIntributes.ListPrice[0].FormattedPrice[0]
                              );
                          }
                          if (
                            results[0].Offers[0].TotalOffers[0] != 0 &&
                            results[0].Offers[0].Offer[0].OfferListing[0]
                              .Price !== undefined
                          ) {
                            console.log(
                              "=>> Price: " +
                                results[0].Offers[0].Offer[0].OfferListing[0]
                                  .Price[0].FormattedPrice[0]
                            );
                            worksheet
                              .cell(i + 2, 6)
                              .string(
                                results[0].Offers[0].Offer[0].OfferListing[0]
                                  .Price[0].FormattedPrice[0]
                              );
                          }
                          if (
                            results[0].Offers[0].TotalOffers[0] != 0 &&
                            results[0].Offers[0].Offer[0].OfferListing[0]
                              .AmountSaved !== undefined
                          ) {
                            console.log(
                              "=>> AmountSaved: " +
                                results[0].Offers[0].Offer[0].OfferListing[0]
                                  .AmountSaved[0].FormattedPrice[0]
                            );
                            worksheet
                              .cell(i + 2, 7)
                              .string(
                                results[0].Offers[0].Offer[0].OfferListing[0]
                                  .AmountSaved[0].FormattedPrice[0]
                              );
                          }

                          if (
                            results[0].Offers[0].TotalOffers[0] != 0 &&
                            results[0].Offers[0].Offer[0].OfferListing[0]
                              .PercentageSaved !== undefined
                          ) {
                            console.log(
                              "=>> PercentageSaved: " +
                                results[0].Offers[0].Offer[0].OfferListing[0]
                                  .PercentageSaved[0]
                            );
                            worksheet
                              .cell(i + 2, 8)
                              .string(
                                results[0].Offers[0].Offer[0].OfferListing[0]
                                  .PercentageSaved[0]
                              );
                          }
                          if (results[0].SalesRank !== undefined) {
                            console.log(
                              "=>> SalesRank: " + results[0].SalesRank[0]
                            );
                            worksheet
                              .cell(i + 2, 9)
                              .string(results[0].SalesRank[0]);
                          }
                          if (ItemIntributes.ProductGroup !== undefined) {
                            console.log(
                              "=>> ProductGroup: " +
                                ItemIntributes.ProductGroup[0]
                            );
                            worksheet
                              .cell(i + 2, 10)
                              .string(ItemIntributes.ProductGroup[0]);
                          }
                          if (ItemIntributes.ProductTypeName !== undefined) {
                            console.log(
                              "=>> ProductTypeName: " +
                                ItemIntributes.ProductTypeName[0]
                            );
                            worksheet
                              .cell(i + 2, 11)
                              .string(ItemIntributes.ProductTypeName[0]);
                          }
                          if (ItemIntributes.UPC !== undefined) {
                            console.log("=>> UPC: " + ItemIntributes.UPC[0]);
                            worksheet
                              .cell(i + 2, 12)
                              .string(ItemIntributes.UPC[0]);
                          }
                          if (
                            results[0].Offers[0].TotalOffers[0] != 0 &&
                            results[0].Offers[0].Offer[0].OfferListing[0]
                              .Availability !== undefined
                          ) {
                            console.log(
                              "=>> Availability: " +
                                results[0].Offers[0].Offer[0].OfferListing[0]
                                  .Availability[0]
                            );
                            worksheet
                              .cell(i + 2, 13)
                              .string(
                                results[0].Offers[0].Offer[0].OfferListing[0]
                                  .Availability[0]
                              );
                          }
                          if (
                            results[0].Offers[0].TotalOffers[0] != 0 &&
                            results[0].Offers[0].Offer[0].OfferListing[0]
                              .IsEligibleForPrime !== undefined
                          ) {
                            console.log(
                              "=>> IsEligibleForPrime: " +
                                results[0].Offers[0].Offer[0].OfferListing[0]
                                  .IsEligibleForPrime[0]
                            );
                            worksheet
                              .cell(i + 2, 14)
                              .string(
                                results[0].Offers[0].Offer[0].OfferListing[0]
                                  .IsEligibleForPrime[0]
                              );
                          }
                          if (ItemIntributes.Manufacturer !== undefined) {
                            console.log(
                              "=>> Manufacturer: " +
                                ItemIntributes.Manufacturer[0]
                            );
                            worksheet
                              .cell(i + 2, 15)
                              .string(ItemIntributes.Manufacturer[0]);
                          }
                          if (ItemIntributes.Brand !== undefined) {
                            console.log(
                              "=>> Brand: " + ItemIntributes.Brand[0]
                            );
                            worksheet
                              .cell(i + 2, 16)
                              .string(ItemIntributes.Brand[0]);
                          }
                          resolve();
                        } catch (e) {
                          console.log(
                            "ERROR IN PROCESSING=>> " + e + " " + formatted
                          );
                          worksheet
                            .cell(i + 2, 19)
                            .string(
                              "Error Occur While Processing This Product " +
                                formatted
                            );
                          resolve();
                        }
                      }
                    }
                  );
                });
              }
            })
            .catch(err => {
              console.error(err);
              console.error("Retrying in 5 seconds");
              retryValue = el;
              setTimeout(resolve(), 5000);
            });
        }, 100);
      });
    },
    options
  ).then(debug => {
    workbook.write("Amazon_Product_Review.xlsx");
    console.log("Finished", debug);
  });
};

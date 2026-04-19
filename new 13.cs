using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Linq;

public class GenerateAutonumberPlugin : IPlugin
{
    public void Execute(IServiceProvider serviceProvider)
    {
        // 🔧 STEP 0: Setup services
        ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
        IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
        IOrganizationServiceFactory factory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
        IOrganizationService service = factory.CreateOrganizationService(context.UserId);

        tracingService.Trace("🔧 STEP 0: Services initialized");

        try
        {
            // ✅ STEP 1: Ensure this is a Create event with a valid target ID
            if (context.MessageName != "Create" || !(context.OutputParameters["id"] is Guid targetId))
            {
                tracingService.Trace("⚠️ STEP 1: Not a Create message or missing Target ID.");
                return;
            }

            string entityName = context.PrimaryEntityName;
            decimal exchangeRate = 0m;
            tracingService.Trace($"✅ STEP 1: Create event confirmed. Entity: {entityName}, ID: {targetId}");

            // 🔍 STEP 2: Retrieve autonumber configuration
            QueryExpression query = new QueryExpression("cr5fb_autonumber")
            {
                ColumnSet = new ColumnSet("cr5fb_counter", "cr5fb_prefix", "cr5fb_padding", "cr5fb_separator"),
                Criteria = new FilterExpression
                {
                    Conditions =
                    {
                        new ConditionExpression("cr5fb_newcolumn", ConditionOperator.Equal, "Travel Authorization")
                    }
                }
            };

            EntityCollection results = service.RetrieveMultiple(query);
            if (results.Entities.Count == 0)
            {
                tracingService.Trace("❌ STEP 2: No autonumber record found with cr5fb_newcolumn = 'Travel Authorization'.");
                return;
            }

            Entity autoNumber = results.Entities.First();
            int currentCounter = autoNumber.GetAttributeValue<int>("cr5fb_counter");
            string prefix = autoNumber.GetAttributeValue<string>("cr5fb_prefix") ?? "";
            int padding = autoNumber.GetAttributeValue<int>("cr5fb_padding");
            string separator = autoNumber.GetAttributeValue<string>("cr5fb_separator") ?? "";

            tracingService.Trace($"📦 STEP 2: Autonumber found. Counter={currentCounter}, Prefix={prefix}, Padding={padding}");

            // 🔢 STEP 3: Increment the counter
            int newCounter = currentCounter + 1;

            Entity updateAuto = new Entity("cr5fb_autonumber", autoNumber.Id);
            updateAuto["cr5fb_counter"] = newCounter;
            service.Update(updateAuto);

            tracingService.Trace($"✅ STEP 3: Counter incremented and saved. New Counter={newCounter}");

            // 🧮 STEP 4: Calculate padding
            int counterLength = newCounter.ToString().Length;
            int paddingZeros = Math.Max(0, padding - counterLength);
            string paddedZeros = new string('0', paddingZeros);

            tracingService.Trace($"🧮 STEP 4: Counter Length={counterLength}, Zero Padding={paddingZeros}, Padding String='{paddedZeros}'");

            // 🏗️ STEP 5: Construct final code
            string finalCode = prefix + separator + paddedZeros + newCounter.ToString();
            tracingService.Trace($"🏗️ STEP 5: Final Code Generated: {finalCode}");

            // --- NEW/UPDATED STEP 5.1: Exchange-rate lookup using cr5fb_exp_exchangerates ---
            tracingService.Trace("💱 STEP 5.1: Exchange rate lookup (custom table) starting...");

            // Retrieve created record to get currency and createdon (date)
            Entity createdEntity = service.Retrieve(entityName, targetId, new ColumnSet("cr5fb_currency", "createdon"));
            Entity updateTarget = new Entity(entityName, targetId);

            // Determine selected currency (if any)
            EntityReference currencyRef = null;
            if (createdEntity.Contains("cr5fb_currency") && createdEntity["cr5fb_currency"] is EntityReference cref)
            {
                currencyRef = cref;
                tracingService.Trace($"🔗 STEP 5.1: Selected currency found: {currencyRef.Id}");
            }
            else
            {
                tracingService.Trace("ℹ️ STEP 5.1: No cr5fb_currency on created record.");
            }

            // get createdon date-only
            if (!createdEntity.Contains("createdon") || createdEntity["createdon"] == null)
            {
                tracingService.Trace("❌ STEP 5.1: createdon missing, cannot determine date for exchange rate lookup.");
            }
            DateTime createdOnDate = createdEntity.GetAttributeValue<DateTime>("createdon").Date;
            tracingService.Trace($"📅 STEP 5.1: Using CreatedOn date (date only): {createdOnDate.ToShortDateString()}");

            // Get base currency (transactioncurrency where currencytype = 1)
            QueryExpression baseCurrencyQuery = new QueryExpression("transactioncurrency")
            {
                ColumnSet = new ColumnSet("transactioncurrencyid"),
                Criteria = new FilterExpression
                {
                    Conditions =
                    {
                        new ConditionExpression("currencytype", ConditionOperator.Equal, 1)
                    }
                },
                TopCount = 1
            };

            EntityCollection baseCurrencyResults = service.RetrieveMultiple(baseCurrencyQuery);
            if (!baseCurrencyResults.Entities.Any())
            {
                tracingService.Trace("❌ STEP 5.1: No base currency found (currencytype = 1). Skipping exchange rate lookup.");
            }
            else
            {
                Guid baseCurrencyId = baseCurrencyResults.Entities.First().Id;
                tracingService.Trace($"🏦 STEP 5.1: Base Currency ID: {baseCurrencyId}");

                // If no currency was selected, we will set default currency later (STEP 5.3).
                if (currencyRef != null)
                {
                    // If selected currency is the base currency => rate = 1
                    if (currencyRef.Id == baseCurrencyId)
                    {
                        tracingService.Trace("💱 STEP 5.1: Selected currency is base currency. Setting exchange rate = 1.");
                        updateTarget["cr5fb_exchangerate"] = 1m;
                        exchangeRate = 1m;
                    }
                    else
                    {
                        // Query cr5fb_exp_exchangerates:
                        // - basecurrency == baseCurrencyId
                        // - destinationcurrency == currencyRef.Id
                        // - date <= createdOnDate
                        // - order by date desc
                        // - take 1
                        QueryExpression exchangeQuery = new QueryExpression("cr5fb_exp_exchangerates")
                        {
                            ColumnSet = new ColumnSet("cr5fb_exchangerate", "cr5fb_date"),
                            Criteria = new FilterExpression(LogicalOperator.And)
                            {
                                Conditions =
                                {
                                    new ConditionExpression("cr5fb_basecurrency", ConditionOperator.Equal, baseCurrencyId),
                                    new ConditionExpression("cr5fb_destinationcurrency", ConditionOperator.Equal, currencyRef.Id),
                                    new ConditionExpression("cr5fb_date", ConditionOperator.LessEqual, createdOnDate)
                                }
                            },
                            Orders =
                            {
                                new OrderExpression("cr5fb_date", OrderType.Descending)
                            },
                            TopCount = 1
                        };

                        EntityCollection exchangeRates = service.RetrieveMultiple(exchangeQuery);

                        if (exchangeRates.Entities.Count > 0)
                        {
                            Entity selectedRate = exchangeRates.Entities.First();
                            exchangeRate = selectedRate.GetAttributeValue<decimal?>("cr5fb_exchangerate") ?? 0m;
                            tracingService.Trace($"💱 STEP 5.1: Exchange rate found: {exchangeRate} (date: {selectedRate.GetAttributeValue<DateTime?>("cr5fb_date")?.ToShortDateString()})");
                            updateTarget["cr5fb_exchangerate"] = exchangeRate;
                        }
                        else
                        {
                            tracingService.Trace("⚠️ STEP 5.1: No matching exchange-rate record found for the currency/date combination.");
                        }
                    }
                }

                // --- STEP 5.3: If currency blank on created record, set default currency + exchangerate = 1
                if (!createdEntity.Contains("cr5fb_currency") || createdEntity["cr5fb_currency"] == null)
                {
                    tracingService.Trace("💲 STEP 5.3: cr5fb_currency is blank. Setting default base currency and exchangerate=1.");

                    Entity defaultCurrency = baseCurrencyResults.Entities.First();
                    updateTarget["cr5fb_currency"] = new EntityReference("transactioncurrency", defaultCurrency.Id);
                    updateTarget["cr5fb_exchangerate"] = 1m;
                    exchangeRate = 1m;
                }
                else
                {
                    tracingService.Trace("💲 STEP 5.3: cr5fb_currency already set on created record; skipping default currency assignment.");
                }
            }

            // 👤 STEP 5.2: Attempting to find employee record for current user.
            tracingService.Trace("👤 STEP 5.2: Attempting to find employee record for current user.");

            QueryExpression employeeQuery = new QueryExpression("cr5fb_employee")
            {
                ColumnSet = new ColumnSet("cr5fb_employeeid"),
                Criteria = new FilterExpression
                {
                    Conditions =
                    {
                        new ConditionExpression("cr5fb_user", ConditionOperator.Equal, context.InitiatingUserId)
                    }
                }
            };

            EntityCollection employeeResults = service.RetrieveMultiple(employeeQuery);

            if (employeeResults.Entities.Count > 0)
            {
                Entity employee = employeeResults.Entities.First();
                tracingService.Trace($"👤 STEP 5.2: Found employee record. ID={employee.Id}");

                updateTarget["cr5fb_employee_requestor"] = new EntityReference("cr5fb_employee", employee.Id);
            }
            else
            {
                tracingService.Trace("⚠️ STEP 5.2: No employee record found for current user.");
            }

            // 📝 STEP 6: Update the created record (finalCode + any currency/exchange/employee set above)
            updateTarget["cr5fb_newcolumn"] = finalCode; // Change this field name to match your real target field
            service.Update(updateTarget);

            tracingService.Trace("✅ STEP 6: Target record updated with the generated code and exchange/currency values.");

            // 🗓️ STEP 6.1: Calculate perdiem days and create perdiem lines
            tracingService.Trace("🗓️ STEP 6.1: Calculating perdiem days...");

            // Retrieve start and end date
            Entity fullEntity = service.Retrieve(entityName, targetId, new ColumnSet("cr5fb_startdate", "cr5fb_enddate", "cr5fb_calculateperdiem", "cr5fb_internationaltravel", "cr5fb_ratedays", "cr5fb_perdiemlumpsump", 
                "cr5fb_country", "cr5fb_city", "cr5fb_departureoneventstartdate", "cr5fb_arrivaloneventenddate"));
            var optionValue = fullEntity.GetAttributeValue<OptionSetValue>("cr5fb_calculateperdiem")?.Value ?? 0;
            if (optionValue == 1)
            {
                if (fullEntity.Contains("cr5fb_startdate") && fullEntity.Contains("cr5fb_enddate"))
                {
                    DateTime startDate = fullEntity.GetAttributeValue<DateTime>("cr5fb_startdate").Date;
                    DateTime endDate = fullEntity.GetAttributeValue<DateTime>("cr5fb_enddate").Date;

                    if (endDate < startDate)
                    {
                        tracingService.Trace("⚠️ STEP 6.1: End date is earlier than start date. Skipping perdiem calculation.");
                    }
                    else
                    {
                        // Check international travel flag
                        bool isInternational = fullEntity.GetAttributeValue<bool>("cr5fb_internationaltravel");
                        bool departonstartdate = fullEntity.GetAttributeValue<bool>("cr5fb_departureoneventstartdate");
                        bool arriveonenddate = fullEntity.GetAttributeValue<bool>("cr5fb_arrivaloneventenddate");
                        int totalDays = 0;
                        if (isInternational)
                        {
                            if (departonstartdate)
                            {
                                totalDays = (endDate - startDate).Days + 1;
                            }
                            else
                            {
                                totalDays = (endDate - startDate).Days + 2;
                            }
                            tracingService.Trace("🌍 STEP 6.1: International travel = YES");
                        }
                        else
                        {
                            if (arriveonenddate)
                            {
                                totalDays = (endDate - startDate).Days + 2;
                            }
                            else
                            {
                                totalDays = (endDate - startDate).Days + 3;
                            }
                            tracingService.Trace("🏠 STEP 6.1: International travel = NO");
                        }
                        tracingService.Trace($"🗓️ STEP 6.1: Perdiem days calculated = {totalDays}");

                        EntityReference countryRef = fullEntity.GetAttributeValue<EntityReference>("cr5fb_country");
                        EntityReference cityRef = fullEntity.GetAttributeValue<EntityReference>("cr5fb_city");
                        decimal rateValue = 0m;

                        if (countryRef != null && cityRef != null)
                        {
                            tracingService.Trace($"🔍 STEP 6.1a: Looking for perdiem rate. Country={countryRef.Id}, City={cityRef.Id}, Date={startDate}");

                            QueryExpression rateQuery = new QueryExpression("cr5fb_hr_perdiemrate")
                            {
                                ColumnSet = new ColumnSet("cr5fb_60days", "cr5fb_60daysup"),
                                Criteria =
                                {
                                    Conditions =
                                    {
                                        new ConditionExpression("cr5fb_country", ConditionOperator.Equal, countryRef.Id),
                                        new ConditionExpression("cr5fb_city", ConditionOperator.Equal, cityRef.Id),
                                        new ConditionExpression("cr5fb_effectivestartdate", ConditionOperator.LessEqual, startDate),
                                        new ConditionExpression("cr5fb_effectiveenddate", ConditionOperator.GreaterEqual, startDate)
                                    }
                                },
                                TopCount = 1
                            };

                            EntityCollection rateResults = service.RetrieveMultiple(rateQuery);

                            if (rateResults.Entities.Count > 0)
                            {
                                Entity rateRecord = rateResults.Entities.First();

                                if (totalDays >= 1 && totalDays <= 60)
                                {
                                    rateValue = (rateRecord.GetAttributeValue<decimal?>("cr5fb_60days") ?? 0m) * exchangeRate;
                                    tracingService.Trace($"💰 STEP 6.1a: Using <=60 days rate. Value={rateValue}");
                                }
                                else
                                {
                                    rateValue = (rateRecord.GetAttributeValue<decimal?>("cr5fb_60daysup") ?? 0m) * exchangeRate;
                                    tracingService.Trace($"💰 STEP 6.1a: Using >60 days rate. Value={rateValue}");
                                }

                                // Save to parent record
                                Entity updateRate = new Entity(entityName, targetId);
                                updateRate["cr5fb_ratedays"] = rateValue;
                                service.Update(updateRate);

                                tracingService.Trace("✅ STEP 6.1a: Rate per days updated on parent record.");
                            }
                            else
                            {
                                tracingService.Trace("⚠️ STEP 6.1a: No perdiem rate record found.");
                            }
                        }
                        else
                        {
                            tracingService.Trace("⚠️ STEP 6.1a: Country or City missing. Cannot lookup perdiem rate.");
                        }

                        // Update parent record with totalDays
                        Entity updateDays = new Entity(entityName, targetId);
                        updateDays["cr5fb_perdiemdays"] = totalDays;
                        service.Update(updateDays);

                        tracingService.Trace("✅ STEP 6.1: Perdiem days updated on parent record.");

                        // Create child records in cr5fb_perdiemline
                        var perdiemLumpsump = fullEntity.GetAttributeValue<OptionSetValue>("cr5fb_perdiemlumpsump")?.Value ?? 0;
                        if (perdiemLumpsump == 1)
                        {
                            Entity perdiemLine = new Entity("cr5fb_perdiemline");
                            perdiemLine["cr5fb_day"] = 1; // sequential day number
                            perdiemLine["cr5fb_travelauthorizationnumber"] = new EntityReference(entityName, targetId); // assuming relationship field is cr5fb_parent

                            service.Create(perdiemLine);
                            tracingService.Trace("💰 STEP 6.2: Perdiem type = Lump-sum. Skipping line item creation.");
                        }
                        else
                        {
                            DateTime currentDate = startDate;
                            if (!departonstartdate)
                            {
                                currentDate = currentDate.AddDays(-1);
                            }
                            for (int i = 1; i <= totalDays; i++)
                            {
                                Entity perdiemLine = new Entity("cr5fb_perdiemline");
                                perdiemLine["cr5fb_day"] = i; // sequential day number
                                perdiemLine["cr5fb_rate"] = rateValue; // rate per days from parent 
                                perdiemLine["cr5fb_travelauthorizationnumber"] = new EntityReference(entityName, targetId); // assuming relationship field is cr5fb_parent

                                bool isWeekend = (currentDate.DayOfWeek == DayOfWeek.Saturday || currentDate.DayOfWeek == DayOfWeek.Sunday);
                                perdiemLine["cr786_weekend"] = isWeekend;
                                perdiemLine["cr786_date"] = currentDate;
                                tracingService.Trace($"📅 STEP 6.2: {currentDate:yyyy-MM-dd} | Weekend={isWeekend}");

                                service.Create(perdiemLine);
                                currentDate = currentDate.AddDays(1);
                                tracingService.Trace($"📌 STEP 6.2: Perdiem line created for day {i}");
                            }
                            tracingService.Trace("🍽️ STEP 6.2: Perdiem type = Daily rate. Creating line items.");
                        }
                        tracingService.Trace("✅ STEP 6.3: All perdiem line records created.");
                    }
                }
                else
                {
                    tracingService.Trace("⚠️ STEP 6.1: StartDate or EndDate missing. Cannot calculate perdiem days.");
                }
            }
            else
            {
                tracingService.Trace("ℹ️ STEP 6.0: Calculate Perdiem = NO. Skipping perdiem calculation.");
            }

            // 📝 STEP 7: Create Timeline Note (Annotation)
            tracingService.Trace("📝 STEP 7: Creating timeline note.");

            Entity note = new Entity("annotation");
            note["subject"] = "Travel Authorization Created";
            note["notetext"] = "Travel Authorization Has Been Created";
            note["objectid"] = new EntityReference(entityName, targetId); // Link to current record

            service.Create(note);

            tracingService.Trace("✅ STEP 7: Timeline note created successfully.");
        }
        catch (Exception ex)
        {
            tracingService.Trace("❌ STEP X: Exception occurred: " + ex.ToString());
            throw new InvalidPluginExecutionException("Error in GenerateAutonumberPlugin", ex);
        }
    }
}

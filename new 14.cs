using System;
using System.Linq;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace TravelAuthorization.Plugins
{
    public class RecalculatePerDiemPlugin : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            try
            {
                if (context.MessageName.ToLower() != "update")
                    return;

                // Target contains updated fields
                Entity target = (Entity)context.InputParameters["Target"];

                // Only proceed if one of the relevant fields changed
                if (!target.Contains("cr5fb_startdate") &&
                    !target.Contains("cr5fb_enddate") &&
                    !target.Contains("cr5fb_perdiemlumpsump") &&
                    !target.Contains("cr5fb_departureoneventstartdate") &&
                    !target.Contains("cr5fb_arrivaloneventenddate") &&
                    !target.Contains("cr5fb_internationaltravel"))
                {
                    tracingService.Trace("⚠️ No change in StartDate, EndDate, or PerdiemLumpsump. Skipping plugin.");
                    return;
                }

                // Get full record with both dates
                Guid recordId = target.Id;
                string entityName = context.PrimaryEntityName;
                Entity fullEntity = service.Retrieve(entityName, recordId, new ColumnSet("cr5fb_startdate", "cr5fb_enddate", "cr5fb_calculateperdiem", "cr5fb_internationaltravel", "cr5fb_ratedays", "cr5fb_perdiemlumpsump", "cr5fb_country", "cr5fb_city", 
                    "cr5fb_exchangerate", "cr5fb_departureoneventstartdate", "cr5fb_arrivaloneventenddate"));

                decimal exchangeRate = (fullEntity.GetAttributeValue<decimal?>("cr5fb_exchangerate") ?? 0m);

                if (!fullEntity.Contains("cr5fb_startdate") || !fullEntity.Contains("cr5fb_enddate") || !fullEntity.Contains("cr5fb_perdiemlumpsump"))
                {
                    tracingService.Trace("⚠️ StartDate or EndDate missing. Skipping calculation.");
                    return;
                }

                // 🗑️ Delete all existing PerDiem lines
                QueryExpression q = new QueryExpression("cr5fb_perdiemline");
                q.ColumnSet = new ColumnSet(false);
                q.Criteria.AddCondition("cr5fb_travelauthorizationnumber", ConditionOperator.Equal, recordId);

                var existingLines = service.RetrieveMultiple(q);
                foreach (var line in existingLines.Entities)
                {
                    service.Delete("cr5fb_perdiemline", line.Id);
                }
                tracingService.Trace($"🗑️ Deleted {existingLines.Entities.Count} existing PerDiem lines.");

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
                            Entity updateRate = new Entity(entityName, recordId);
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
                    Entity updateDays = new Entity(entityName, recordId);
                    updateDays["cr5fb_perdiemdays"] = totalDays;
                    service.Update(updateDays);

                    tracingService.Trace("✅ STEP 6.1: Perdiem days updated on parent record.");

                    // Create child records in cr5fb_perdiemline
                    var perdiemLumpsump = fullEntity.GetAttributeValue<OptionSetValue>("cr5fb_perdiemlumpsump")?.Value ?? 0;
                    if (perdiemLumpsump == 1)
                    {
                        Entity perdiemLine = new Entity("cr5fb_perdiemline");
                        perdiemLine["cr5fb_day"] = 1; // sequential day number
                        perdiemLine["cr5fb_travelauthorizationnumber"] = new EntityReference(entityName, recordId); // assuming relationship field is cr5fb_parent

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
                            perdiemLine["cr5fb_travelauthorizationnumber"] = new EntityReference(entityName, recordId); // assuming relationship field is cr5fb_parent

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
            catch (Exception ex)
            {
                tracingService.Trace("❌ Exception: " + ex.ToString());
                throw new InvalidPluginExecutionException("Recalculate PerDiem plugin failed: " + ex.Message);
            }
        }
    }
}

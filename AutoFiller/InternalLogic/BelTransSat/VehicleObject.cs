
using System.Text.Json.Serialization;

namespace ExcelParser.BelTransSat
{
    public class VehicleObject
    {
        public VehicleObject()
        {
            FuelInList = new List<Refuel>();
        }
        
        [JsonPropertyName("object_id")]
        public string ObjectId { get; set; }
        [JsonPropertyName("object_name")]
        public string ObjectName { get; set; }
        [JsonPropertyName("object_uid")]
        public string ObjectUid { get; set; }
        [JsonPropertyName("distance_gps")]
        public double? DistanceGps { get; set; }
        [JsonPropertyName("distance_can")] 
        public double? DistanceCan { get; set; }
        [JsonPropertyName("run_time")]
        public double? RunTime { get; set; }
        [JsonPropertyName("run_time_str")]
        public string RunTimeStr { get; set; }
        [JsonPropertyName("stop_time")]
        public double? StopTime { get; set; }
        [JsonPropertyName("stop_time_str")]
        public string StopTimeStr { get; set; }
        [JsonPropertyName("max_speed")]
        public double? MaxSpeed { get; set; }
        [JsonPropertyName("fuel_can")]
        public double? FuelCan { get; set; }
        [JsonPropertyName("fuel_dut")]
        public double? FuelDut { get; set; }
        
        [JsonPropertyName("fuel_in_list")]
        public List<Refuel> FuelInList { get; set; }

        [JsonPropertyName("fuel_dut_start")]
        public double? FuelDutStart { get; set; }
        [JsonPropertyName("fuel_dut_finish")]
        public double? FuelDutFinish { get; set; }
        [JsonPropertyName("odom_start")]
        public double? OdomStart { get; set; }
        [JsonPropertyName("s_odo_dt")]
        public string SOdoDt { get; set; }
        [JsonPropertyName("odom_finish")]
        public double? OdomFinish { get; set; }
        [JsonPropertyName("e_odo_dt")]
        public string EOdoDt { get; set; }
        [JsonPropertyName("avg_speed_gps")]
        public double? AvgSpeedGps { get; set; }
        [JsonPropertyName("avg_speed_can")]
        public double? AvgSpeedCan { get; set; }
        [JsonPropertyName("avg_fuel_can")]
        public double? AvgFuelCan { get; set; }
        [JsonPropertyName("avg_fuel_dut")]
        public double? AvgFuelDut { get; set; }
        [JsonPropertyName("fuel_can_stop")]
        public double? FuelCanStop { get; set; }
        [JsonPropertyName("engine_time_h")]
        public double? EngineTimeH { get; set; }

        public double? GetTravelDistance()
        {
            if (DistanceGps == null && DistanceCan == null)
            {
                return 0f;
            }

            if (DistanceGps == null)
            {
                return DistanceCan;
            }
            
            if (DistanceCan == null)
            {
                return DistanceGps;
            }
            
            return (DistanceGps + DistanceCan) / 2;
        }
        
        public double? GetFuelUsed()
        {
            if (FuelDut == null && FuelCan == null)
            {
                return 0f;
            }

            if (FuelDut == null)
            {
                return FuelCan;
            }
            
            if (FuelCan == null)
            {
                return FuelDut;
            }
            
            return (FuelDut + FuelCan) / 2;
        }

        public double? GetMachineHours()
        {
            if (EngineTimeH == null)
            {
                return 0;
            }
            
            return EngineTimeH;
        }

        public double? GetRefuel()
        {
            double? sum=0;
            
            foreach (var refuel in FuelInList)
            {
                sum += refuel.GetRefuelValue();
            }
            
            return sum;
        }
    }

    public class RootObject
    {
        [JsonPropertyName("root")]
        public ResultWrapper RootWrapper { get; set; }

        public VehicleObject FindWithId(string id)
        {
            foreach (VehicleObject vehicle in RootWrapper.Result.Items)
            {
                if (vehicle.ObjectName == id)
                {
                    return vehicle;
                }
            }

            return new VehicleObject();
        }
    }

    public class ResultWrapper
    {
        [JsonPropertyName("result")]
        public Result Result { get; set; }
    }

    public class Result
    {
        [JsonPropertyName("items")]
        public List<VehicleObject> Items { get; set; }
    }

    public class Refuel
    {
        [JsonPropertyName("dt")]
        public string Dt { get; set; }
        [JsonPropertyName("value")]
        public double? Value { get; set; }
        [JsonPropertyName("lat")]
        public string Lat { get; set; }
        [JsonPropertyName("lon")]
        public string Lon { get; set; }
        [JsonPropertyName("address")]
        public string Address { get; set; }

        public double? GetRefuelValue()
        {
            if (Value == null)
            {
                return 0;
            }
            return Value;
        }
    }
}
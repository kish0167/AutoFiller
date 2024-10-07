using DataTracker.Utility;
using ExcelParser;

namespace DataTracker.Internal
{
    public class VehicleFuelStatistics
    {
        private string _name;

        public string Name
        {
            get => _name;
            set => _name = value;
        }

        private List<double> _refuels;
        private List<double> _travelsDistances;
        private List<double> _theoryConsumptions;

        public List<double> TheoryConsumptions => _theoryConsumptions;

        private double _fuelConsumption;

        public VehicleFuelStatistics()
        {
            _name = " none ";
            _refuels = new List<double>();
            _travelsDistances = new List<double>();
            _theoryConsumptions = new List<double>();
            _fuelConsumption = 0;
        }

        public double FuelConsumption
        {
            get => _fuelConsumption;
            set => _fuelConsumption = value;
        }

        public void AddRefuel(double value)
        {
            _refuels.Add(value);
        }

        public void AddTravel(double value)
        {
            _travelsDistances.Add(value);
        }

        public void PrintAll()
        {
            for (int i = 0; i < _travelsDistances.Count; i++)
            {
                Logger.Log("All data in " + _name);
                Console.WriteLine($"{_travelsDistances[i]}\t{_refuels[i]}");
            }
        }

        public void CalculateTheoryConsumption()
        {
            double consumptionBuf = 0f;
            double volumeBuf = 0;
            double distanceBuf = 0;
            int lastRefuelIndex = 0;

            for (int i = 0; i < _refuels.Count; i++)
            {
                if (_refuels[i] != 0)
                {
                    if (distanceBuf != 0)
                    {
                        consumptionBuf = volumeBuf * 100 / distanceBuf;
                    }

                    for (int j = 0; j < i - lastRefuelIndex; j++)
                    {
                        _theoryConsumptions.Add(consumptionBuf);
                    }

                    lastRefuelIndex = i;
                    volumeBuf = _refuels[i];
                    distanceBuf = _travelsDistances[i];
                }
                else
                {
                    distanceBuf += _travelsDistances[i];
                }
            }
        }
    }
}
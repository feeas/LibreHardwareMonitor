// This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
// If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
// Copyright (C) LibreHardwareMonitor and Contributors.
// Partial Copyright (C) Michael Möller <mmoeller@openhardwaremonitor.org> and Contributors.
// All Rights Reserved.

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using LibreHardwareMonitor.Hardware.Battery;
using LibreHardwareMonitor.Hardware.Controller.AeroCool;
using LibreHardwareMonitor.Hardware.Controller.AquaComputer;
using LibreHardwareMonitor.Hardware.Controller.Heatmaster;
using LibreHardwareMonitor.Hardware.Controller.Nzxt;
using LibreHardwareMonitor.Hardware.Controller.Razer;
using LibreHardwareMonitor.Hardware.Controller.TBalancer;
using LibreHardwareMonitor.Hardware.Cpu;
using LibreHardwareMonitor.Hardware.Gpu;
using LibreHardwareMonitor.Hardware.Memory;
using LibreHardwareMonitor.Hardware.Motherboard;
using LibreHardwareMonitor.Hardware.Network;
using LibreHardwareMonitor.Hardware.Psu.Corsair;
using LibreHardwareMonitor.Hardware.Psu.Msi;
using LibreHardwareMonitor.Hardware.Storage;

namespace LibreHardwareMonitor.Hardware;

/// <summary>
/// Stores all hardware groups and decides which devices should be enabled and updated.
/// </summary>
public class Computer : IComputer
{
    private readonly List<IGroup> _groups = new();
    private readonly object _lock = new();
    private readonly ISettings _settings;

    private bool _batteryEnabled;
    private bool _controllerEnabled;
    private bool _cpuEnabled;
    private bool _gpuEnabled;
    private bool _memoryEnabled;
    private bool _motherboardEnabled;
    private bool _networkEnabled;
    private bool _open;
    private bool _psuEnabled;
    private SMBios _smbios;
    private bool _storageEnabled;

    /// <summary>
    /// Creates a new <see cref="IComputer" /> instance with basic initial <see cref="Settings" />.
    /// </summary>
    public Computer()
    {
        _settings = new Settings();
    }

    /// <summary>
    /// Creates a new <see cref="IComputer" /> instance with additional <see cref="ISettings" />.
    /// </summary>
    /// <param name="settings">Computer settings that will be transferred to each <see cref="IHardware" />.</param>
    public Computer(ISettings settings)
    {
        _settings = settings ?? new Settings();
    }

    /// <inheritdoc />
    public event HardwareEventHandler HardwareAdded;

    /// <inheritdoc />
    public event HardwareEventHandler HardwareRemoved;

    /// <inheritdoc />
    public IList<IHardware> Hardware
    {
        get
        {
            lock (_lock)
            {
                List<IHardware> list = new();

                foreach (IGroup group in _groups)
                    list.AddRange(group.Hardware);

                return list;
            }
        }
    }

    /// <inheritdoc />
    public bool IsBatteryEnabled
    {
        get { return _batteryEnabled; }
        set
        {
            if (_open && value != _batteryEnabled)
            {
                if (value)
                {
                    Add(new BatteryGroup(_settings));
                }
                else
                {
                    RemoveType<BatteryGroup>();
                }
            }

            _batteryEnabled = value;
        }
    }

    /// <inheritdoc />
    public bool IsControllerEnabled
    {
        get { return _controllerEnabled; }
        set
        {
            if (_open && value != _controllerEnabled)
            {
                if (value)
                {
                    Add(new TBalancerGroup(_settings));
                    Add(new HeatmasterGroup(_settings));
                    Add(new AquaComputerGroup(_settings));
                    Add(new AeroCoolGroup(_settings));
                    Add(new NzxtGroup(_settings));
                    Add(new RazerGroup(_settings));
                }
                else
                {
                    RemoveType<TBalancerGroup>();
                    RemoveType<HeatmasterGroup>();
                    RemoveType<AquaComputerGroup>();
                    RemoveType<AeroCoolGroup>();
                    RemoveType<NzxtGroup>();
                    RemoveType<RazerGroup>();
                }
            }

            _controllerEnabled = value;
        }
    }

    /// <inheritdoc />
    public bool IsCpuEnabled
    {
        get { return _cpuEnabled; }
        set
        {
            if (_open && value != _cpuEnabled)
            {
                if (value)
                    Add(new CpuGroup(_settings));
                else
                    RemoveType<CpuGroup>();
            }

            _cpuEnabled = value;
        }
    }

    /// <inheritdoc />
    public bool IsGpuEnabled
    {
        get { return _gpuEnabled; }
        set
        {
            if (_open && value != _gpuEnabled)
            {
                if (value)
                {
                    Add(new AmdGpuGroup(_settings));
                    Add(new NvidiaGroup(_settings));

                    if (_cpuEnabled)
                        Add(new IntelGpuGroup(GetIntelCpus(), _settings));
                }
                else
                {
                    RemoveType<AmdGpuGroup>();
                    RemoveType<NvidiaGroup>();
                    RemoveType<IntelGpuGroup>();
                }
            }

            _gpuEnabled = value;
        }
    }

    /// <inheritdoc />
    public bool IsMemoryEnabled
    {
        get { return _memoryEnabled; }
        set
        {
            if (_open && value != _memoryEnabled)
            {
                if (value)
                    Add(new MemoryGroup(_settings));
                else
                    RemoveType<MemoryGroup>();
            }

            _memoryEnabled = value;
        }
    }

    /// <inheritdoc />
    public bool IsMotherboardEnabled
    {
        get { return _motherboardEnabled; }
        set
        {
            if (_open && value != _motherboardEnabled)
            {
                if (value)
                    Add(new MotherboardGroup(_smbios, _settings));
                else
                    RemoveType<MotherboardGroup>();
            }

            _motherboardEnabled = value;
        }
    }

    /// <inheritdoc />
    public bool IsNetworkEnabled
    {
        get { return _networkEnabled; }
        set
        {
            if (_open && value != _networkEnabled)
            {
                if (value)
                    Add(new NetworkGroup(_settings));
                else
                    RemoveType<NetworkGroup>();
            }

            _networkEnabled = value;
        }
    }

    /// <inheritdoc />
    public bool IsPsuEnabled
    {
        get { return _psuEnabled; }
        set
        {
            if (_open && value != _psuEnabled)
            {
                if (value)
                {
                    Add(new CorsairPsuGroup(_settings));
                    Add(new MsiPsuGroup(_settings));
                }
                else
                {
                    RemoveType<CorsairPsuGroup>();
                    RemoveType<MsiPsuGroup>();
                }
            }

            _psuEnabled = value;
        }
    }

    /// <inheritdoc />
    public bool IsStorageEnabled
    {
        get { return _storageEnabled; }
        set
        {
            if (_open && value != _storageEnabled)
            {
                if (value)
                    Add(new StorageGroup(_settings));
                else
                    RemoveType<StorageGroup>();
            }

            _storageEnabled = value;
        }
    }

    /// <summary>
    /// Contains computer information table read in accordance with <see href="https://www.dmtf.org/standards/smbios">System Management BIOS (SMBIOS) Reference Specification</see>.
    /// </summary>
    public SMBios SMBios
    {
        get
        {
            if (!_open)
                throw new InvalidOperationException("SMBIOS cannot be accessed before opening.");

            return _smbios;
        }
    }

    //// <inheritdoc />
    public string GetReport()
    {
        lock (_lock)
        {
            using StringWriter w = new(CultureInfo.InvariantCulture);

            w.WriteLine();
            w.WriteLine(nameof(LibreHardwareMonitor) + " Report");
            w.WriteLine();

            Version version = typeof(Computer).Assembly.GetName().Version;

            NewSection(w);
            w.Write("Version: ");
            w.WriteLine(version.ToString());
            w.WriteLine();

            NewSection(w);
            w.Write("Common Language Runtime: ");
            w.WriteLine(Environment.Version.ToString());
            w.Write("Operating System: ");
            w.WriteLine(Environment.OSVersion.ToString());
            w.Write("Process Type: ");
            w.WriteLine(IntPtr.Size == 4 ? "32-Bit" : "64-Bit");
            w.WriteLine();

            NewSection(w);
            w.WriteLine("Sensors");
            w.WriteLine();

            foreach (IGroup group in _groups)
            {
                foreach (IHardware hardware in group.Hardware)
                    ReportHardwareSensorTree(hardware, w, string.Empty);
            }

            w.WriteLine();

            NewSection(w);
            w.WriteLine("Parameters");
            w.WriteLine();

            foreach (IGroup group in _groups)
            {
                foreach (IHardware hardware in group.Hardware)
                    ReportHardwareParameterTree(hardware, w, string.Empty);
            }

            w.WriteLine();

            foreach (IGroup group in _groups)
            {
                string report = group.GetReport();
                if (!string.IsNullOrEmpty(report))
                {
                    NewSection(w);
                    w.Write(report);
                }

                foreach (IHardware hardware in group.Hardware)
                    ReportHardware(hardware, w);
            }

            return w.ToString();
        }
    }

    /// <summary>
    /// Triggers the <see cref="IVisitor.VisitComputer" /> method for the given observer.
    /// </summary>
    /// <param name="visitor">Observer who call to devices.</param>
    public void Accept(IVisitor visitor)
    {
        if (visitor == null)
            throw new ArgumentNullException(nameof(visitor));

        visitor.VisitComputer(this);
    }

    /// <summary>
    /// Triggers the <see cref="IElement.Accept" /> method with the given visitor for each device in each group.
    /// </summary>
    /// <param name="visitor">Observer who call to devices.</param>
    public void Traverse(IVisitor visitor)
    {
        lock (_lock)
        {
            // Use a for-loop instead of foreach to avoid a collection modified exception after sleep, even though everything is under a lock.
            for (int i = 0; i < _groups.Count; i++)
            {
                IGroup group = _groups[i];

                for (int j = 0; j < group.Hardware.Count; j++)
                    group.Hardware[j].Accept(visitor);
            }
        }
    }

    private void HardwareAddedEvent(IHardware hardware)
    {
        HardwareAdded?.Invoke(hardware);
    }

    private void HardwareRemovedEvent(IHardware hardware)
    {
        HardwareRemoved?.Invoke(hardware);
    }

    private void Add(IGroup group)
    {
        if (group == null)
            return;

        lock (_lock)
        {
            if (_groups.Contains(group))
                return;

            _groups.Add(group);

            if (group is IHardwareChanged hardwareChanged)
            {
                hardwareChanged.HardwareAdded += HardwareAddedEvent;
                hardwareChanged.HardwareRemoved += HardwareRemovedEvent;
            }
        }

        if (HardwareAdded != null)
        {
            foreach (IHardware hardware in group.Hardware)
                HardwareAdded(hardware);
        }
    }

    private void Remove(IGroup group)
    {
        lock (_lock)
        {
            if (!_groups.Contains(group))
                return;

            _groups.Remove(group);

            if (group is IHardwareChanged hardwareChanged)
            {
                hardwareChanged.HardwareAdded -= HardwareAddedEvent;
                hardwareChanged.HardwareRemoved -= HardwareRemovedEvent;
            }
        }

        if (HardwareRemoved != null)
        {
            foreach (IHardware hardware in group.Hardware)
                HardwareRemoved(hardware);
        }

        group.Close();
    }

    private void RemoveType<T>() where T : IGroup
    {
        List<T> list = [];

        lock (_lock)
        {
            foreach (IGroup group in _groups)
            {
                if (group is T t)
                    list.Add(t);
            }
        }

        foreach (T group in list)
            Remove(group);
    }

    /// <summary>
    /// If hasn't been opened before, opens <see cref="SMBios" />, <see cref="OpCode" /> and triggers the private <see cref="AddGroups" /> method depending on which categories are
    /// enabled.
    /// </summary>
    public void Open()
    {
        if (_open)
            return;

        _smbios = new SMBios();

        Mutexes.Open();
        OpCode.Open();

        AddGroups();

        _open = true;
    }

    private void AddGroups()
    {
        if (_motherboardEnabled)
            Add(new MotherboardGroup(_smbios, _settings));

        if (_cpuEnabled)
            Add(new CpuGroup(_settings));

        if (_memoryEnabled)
            Add(new MemoryGroup(_settings));

        if (_gpuEnabled)
        {
            Add(new AmdGpuGroup(_settings));
            Add(new NvidiaGroup(_settings));

            if (_cpuEnabled)
                Add(new IntelGpuGroup(GetIntelCpus(), _settings));
        }

        if (_controllerEnabled)
        {
            Add(new TBalancerGroup(_settings));
            Add(new HeatmasterGroup(_settings));
            Add(new AquaComputerGroup(_settings));
            Add(new AeroCoolGroup(_settings));
            Add(new NzxtGroup(_settings));
            Add(new RazerGroup(_settings));
        }

        if (_storageEnabled)
            Add(new StorageGroup(_settings));

        if (_networkEnabled)
            Add(new NetworkGroup(_settings));

        if (_psuEnabled)
        {
            Add(new CorsairPsuGroup(_settings));
            Add(new MsiPsuGroup(_settings));
        }

        if (_batteryEnabled)
            Add(new BatteryGroup(_settings));
    }

    private static void NewSection(TextWriter writer)
    {
        for (int i = 0; i < 8; i++)
            writer.Write("----------");

        writer.WriteLine();
        writer.WriteLine();
    }

    private static int CompareSensor(ISensor a, ISensor b)
    {
        int c = a.SensorType.CompareTo(b.SensorType);
        if (c == 0)
            return a.Index.CompareTo(b.Index);

        return c;
    }

    private static void ReportHardwareSensorTree(IHardware hardware, TextWriter w, string space)
    {
        w.WriteLine("{0}|", space);
        w.WriteLine("{0}+- {1} ({2})", space, hardware.Name, hardware.Identifier);

        ISensor[] sensors = hardware.Sensors;
        Array.Sort(sensors, CompareSensor);

        foreach (ISensor sensor in sensors)
            w.WriteLine("{0}|  +- {1,-14} : {2,8:G6} {3,8:G6} {4,8:G6} ({5})", space, sensor.Name, sensor.Value, sensor.Min, sensor.Max, sensor.Identifier);

        foreach (IHardware subHardware in hardware.SubHardware)
            ReportHardwareSensorTree(subHardware, w, "|  ");
    }

    private static void ReportHardwareParameterTree(IHardware hardware, TextWriter w, string space)
    {
        w.WriteLine("{0}|", space);
        w.WriteLine("{0}+- {1} ({2})", space, hardware.Name, hardware.Identifier);

        ISensor[] sensors = hardware.Sensors;
        Array.Sort(sensors, CompareSensor);

        foreach (ISensor sensor in sensors)
        {
            string innerSpace = space + "|  ";
            if (sensor.Parameters.Count > 0)
            {
                w.WriteLine("{0}|", innerSpace);
                w.WriteLine("{0}+- {1} ({2})", innerSpace, sensor.Name, sensor.Identifier);

                foreach (IParameter parameter in sensor.Parameters)
                {
                    string innerInnerSpace = innerSpace + "|  ";
                    w.WriteLine("{0}+- {1} : {2}", innerInnerSpace, parameter.Name, string.Format(CultureInfo.InvariantCulture, "{0} : {1}", parameter.DefaultValue, parameter.Value));
                }
            }
        }

        foreach (IHardware subHardware in hardware.SubHardware)
            ReportHardwareParameterTree(subHardware, w, "|  ");
    }

    private static void ReportHardware(IHardware hardware, TextWriter w)
    {
        string hardwareReport = hardware.GetReport();
        if (!string.IsNullOrEmpty(hardwareReport))
        {
            NewSection(w);
            w.Write(hardwareReport);
        }

        foreach (IHardware subHardware in hardware.SubHardware)
            ReportHardware(subHardware, w);
    }

    /// <summary>
    /// If opened before, removes all <see cref="IGroup" /> and triggers <see cref="OpCode.Close" />.
    /// </summary>
    public void Close()
    {
        if (!_open)
            return;

        lock (_lock)
        {
            while (_groups.Count > 0)
            {
                IGroup group = _groups[_groups.Count - 1];
                Remove(group);
            }
        }

        OpCode.Close();
        Mutexes.Close();

        _smbios = null;
        _open = false;
    }

    /// <summary>
    /// If opened before, removes all <see cref="IGroup" /> and recreates it.
    /// </summary>
    public void Reset()
    {
        if (!_open)
            return;

        RemoveGroups();
        AddGroups();
    }

    private void RemoveGroups()
    {
        lock (_lock)
        {
            while (_groups.Count > 0)
            {
                IGroup group = _groups[_groups.Count - 1];
                Remove(group);
            }
        }
    }

    private List<IntelCpu> GetIntelCpus()
    {
        // Create a temporary cpu group if one has not been added.
        lock (_lock)
        {
            IGroup cpuGroup = _groups.Find(x => x is CpuGroup) ?? new CpuGroup(_settings);
            return cpuGroup.Hardware.Select(x => x as IntelCpu).ToList();
        }
    }

    /// <summary>
    /// <see cref="Computer" /> specific additional settings passed to its <see cref="IHardware" />.
    /// </summary>
    private class Settings : ISettings
    {
        public bool Contains(string name)
        {
            return false;
        }

        public void SetValue(string name, string value)
        { }

        public string GetValue(string name, string value)
        {
            return value;
        }

        public void Remove(string name)
        { }
    }


    private int _CPUPowerIndex;
    private int _CPUClockMaxNumber;
    private int _CPUClockMinIndex;
    private int _CPUTemperatureIndex;
    private int _CPUUtilizationIndex;
    private int _GPUNameIndex;
    // private int _iGPUPowerIndex;
    private IComputer _Computer;
    private string _D3DDisplayDeviceIdentifier;
    private long _gpuNodeUsagePrevValue;
    private DateTime _gpuNodeUsagePrevTick;
    private D3DDisplayDevice.D3DDeviceInfo _deviceInfo;




    public void InitAhk(IComputer Computer)
    {
        _Computer = Computer;
    }


    private int _GPUClockIndex;
    private int _GPUTemperatureIndex;
    private int _GPUPowerIndex;
    private int _VramUsedIndex;
    private int _VramFreeIndex;
    private int _VramTotalIndex;
    private float _VramTotal;
    private int _GPUUtilizationIndex;



    private int _StorageTemperatureIndex;
    private float[] _AllStorageData;
    private int _StorageReadIndex;
    private int _StorageWriteIndex;
    public int InitGetAllStorageData()
    {
        _AllStorageData = new float[3];
        GetStorageTemperatureIndex();
        GetStorageReadIndex();
        GetStorageWriteIndex();

        return 0;
    }
    public float[] GetAllStorageData(int StorageIndex)
    {
        _Computer.Hardware[StorageIndex].Update();
        _AllStorageData[0] = (float)_Computer.Hardware[StorageIndex].Sensors[_StorageTemperatureIndex].Value;
        _AllStorageData[1] = (float)_Computer.Hardware[StorageIndex].Sensors[_StorageReadIndex].Value;
        _AllStorageData[2] = (float)_Computer.Hardware[StorageIndex].Sensors[_StorageWriteIndex].Value;
        return _AllStorageData;
    }

    public int GetStorageTemperatureIndex()
    {
        foreach (IGroup group in _groups)
        {
            foreach (IHardware hardware in group.Hardware)
            {
                //w.WriteLine(hardware.Identifier);
                if (hardware.HardwareType == HardwareType.Storage)
                {
                    for (int j = 0; j < hardware.Sensors.Length; j++)
                    {
                        //找到温度传感器
                        if (hardware.Sensors[j].SensorType == SensorType.Temperature)
                        {
                            if (hardware.Sensors[j].Name == "Temperature")
                            {
                                _StorageTemperatureIndex = j;
                                return j;
                            }
                        }
                    }
                }
            }
        }
        return -1;
    }

    public int GetStorageReadIndex()
    {
        foreach (IGroup group in _groups)
        {
            foreach (IHardware hardware in group.Hardware)
            {
                //w.WriteLine(hardware.Identifier);
                if (hardware.HardwareType == HardwareType.Storage)
                {
                    for (int j = 0; j < hardware.Sensors.Length; j++)
                    {
                        //找到温度传感器
                        if (hardware.Sensors[j].SensorType == SensorType.Throughput)
                        {
                            if (hardware.Sensors[j].Name == "Read Rate")
                            {
                                _StorageReadIndex = j;
                                return j;
                            }
                        }
                    }
                }
            }
        }
        return -1;
    }

    public int GetStorageWriteIndex()
    {
        foreach (IGroup group in _groups)
        {
            foreach (IHardware hardware in group.Hardware)
            {
                //w.WriteLine(hardware.Identifier);
                if (hardware.HardwareType == HardwareType.Storage)
                {
                    for (int j = 0; j < hardware.Sensors.Length; j++)
                    {
                        //找到温度传感器
                        if (hardware.Sensors[j].SensorType == SensorType.Throughput)
                        {
                            if (hardware.Sensors[j].Name == "Write Rate")
                            {
                                _StorageWriteIndex = j;
                                return j;
                            }
                        }
                    }
                }
            }
        }
        return -1;
    }


    // private float[] _AllCpu_IGPU_Data;
    public int InitGetAllCpuData(IComputer Computer)
    {
        // _AllCpu_IGPU_Data = new float[4];
        // _AllCpu_IGPU_Data = new float[5];
        _CPUClockMaxNumber = GetCPUClockIndex();
        _CPUTemperatureIndex = GetCPUTemperatureIndex();
        _CPUPowerIndex = GetCPUPowerIndex();
        _CPUUtilizationIndex = GetCPUUtilizationIndex();
        // GetiGPUPowerIndex();
        _Computer = Computer;
        return 0;
    }

    private float[] _AllCpu_GPU_Data;
    public int InitGetAllGpuData()
    {
        _AllCpu_GPU_Data = new float[9];
        _GPUPowerIndex = GetGPUPowerIndex();
        _GPUUtilizationIndex = GetGPUUtilizationIndex();
        // GetVramUsedIndex();
        _VramFreeIndex = GetVramFreeIndex();
        // GetVramTotalIndex();
        _GPUClockIndex = GetGPUClockIndex();
        _GPUTemperatureIndex = GetGPUTemperatureIndex();
        // _VramTotal =(int)_Computer.Hardware[1].Sensors[_VramTotalIndex].Value;
        return 0;
    }

    private float[] _AllCpu_IGPU_Data;
    public int InitGetAlliGpuData()
    {
        _AllCpu_IGPU_Data = new float[7];
        _Computer.Hardware[1].Update();

        _GPUPowerIndex = GetGPUPowerIndex();
        _GPUUtilizationIndex = GetGPUUtilizationIndex();
        // GetVramUsedIndex();
        _VramFreeIndex = GetVramFreeIndex();
        // _VramTotal =GetGpuSharedLimit(GetD3DDisplayDeviceIdentifier(0))/1048576;

        return 0;
    }

    public float[] GetAllCpu_IGPU_Data()
    {
        _Computer.Hardware[0].Update();
        _AllCpu_IGPU_Data[0] = GetMaxCPUClock();
        _AllCpu_IGPU_Data[1] = (float)_Computer.Hardware[0].Sensors[_CPUTemperatureIndex].Value;
        _AllCpu_IGPU_Data[2] = (float)_Computer.Hardware[0].Sensors[_CPUPowerIndex].Value;
        _AllCpu_IGPU_Data[3] = (float)_Computer.Hardware[0].Sensors[_CPUUtilizationIndex].Value;
        _Computer.Hardware[1].Update();
        _AllCpu_IGPU_Data[4] = (float)_Computer.Hardware[1].Sensors[_GPUPowerIndex].Value;
        _AllCpu_IGPU_Data[5] = (float)_Computer.Hardware[1].Sensors[_GPUUtilizationIndex].Value;
        _AllCpu_IGPU_Data[6] = (float)_Computer.Hardware[1].Sensors[_VramFreeIndex].Value;
        // _AllCpu_IGPU_Data[6] = (float)(_VramTotal-_Computer.Hardware[1].Sensors[_VramUsedIndex].Value);
        // if (_iGPUPowerIndex != -1)
        // {
        //     _AllCpu_IGPU_Data[4] = (float)_Computer.Hardware[0].Sensors[_iGPUPowerIndex].Value;
        // }


        return _AllCpu_IGPU_Data;
    }

    public float GetMaxCPUClock()
    {
        float MaxClock = 0;
        for (int i = 0; i < _CPUClockMaxNumber; i++)
        {
            if (MaxClock < _Computer.Hardware[0].Sensors[_CPUClockMinIndex + i].Value)
            {
                MaxClock = (float)_Computer.Hardware[0].Sensors[_CPUClockMinIndex + i].Value;
            }
        }
        return MaxClock;
    }


    private float[] _AllCpu_GPU_Data_Safe;
    public int InitGetAllGpuData_Safe()
    {

        // if(_CPUTemperatureIndex==-1){
        //     _CPUTemperatureIndex=GetCPUTemperatureIndex_Safe();
        // }
        _AllCpu_GPU_Data_Safe = new float[9];
        _GPUPowerIndex = GetGPUPowerIndex();
        _GPUUtilizationIndex = GetGPUUtilizationIndex();
        // GetVramUsedIndex();
        _VramFreeIndex = GetVramFreeIndex();
        // GetVramTotalIndex();
        _GPUClockIndex = GetGPUClockIndex();
        _GPUTemperatureIndex = GetGPUTemperatureIndex();
        return 0;
    }
    public float[] GetAllCpu_GPU_Data_Safe()
    {
        _Computer.Hardware[0].Update();
        if (_CPUClockMinIndex == -1)
        {
            _AllCpu_GPU_Data_Safe[0] = -1000;
        }
        else
        {
            _AllCpu_GPU_Data_Safe[0] = GetMaxCPUClock();
        }

        if (_CPUTemperatureIndex == -1)
        {
            _AllCpu_GPU_Data_Safe[1] = -1;
        }
        else
        {
            _AllCpu_GPU_Data_Safe[1] = (float)_Computer.Hardware[0].Sensors[_CPUTemperatureIndex].Value;
        }

        if (_CPUPowerIndex == -1)
        {
            _AllCpu_GPU_Data_Safe[2] = -1;
        }
        else
        {
            _AllCpu_GPU_Data_Safe[2] = (float)_Computer.Hardware[0].Sensors[_CPUPowerIndex].Value;
        }

        if (_CPUUtilizationIndex == -1)
        {
            _AllCpu_GPU_Data_Safe[3] = -1;
        }
        else
        {
            _AllCpu_GPU_Data_Safe[3] = (float)_Computer.Hardware[0].Sensors[_CPUUtilizationIndex].Value;
        }


        if (_GPUNameIndex == -1)
        {
            for (int i = 4; i < 9; i++)
            {
                _AllCpu_GPU_Data_Safe[i] = -1;
            }
            _AllCpu_GPU_Data_Safe[6] = -1024;
            _AllCpu_GPU_Data_Safe[7] = -1000;
        }
        else
        {
            _Computer.Hardware[1].Update();
            if (_GPUPowerIndex == -1)
            {
                _AllCpu_GPU_Data_Safe[4] = -1;
            }
            else
            {
                _AllCpu_GPU_Data_Safe[4] = (float)_Computer.Hardware[1].Sensors[_GPUPowerIndex].Value;
            }
            if (_GPUUtilizationIndex == -1)
            {
                _AllCpu_GPU_Data_Safe[5] = -1;
            }
            else
            {
                _AllCpu_GPU_Data_Safe[5] = (float)_Computer.Hardware[1].Sensors[_GPUUtilizationIndex].Value;
            }
            if (_VramFreeIndex == -1)
            {
                _AllCpu_GPU_Data_Safe[6] = -1024;
            }
            else
            {
                _AllCpu_GPU_Data_Safe[6] = (float)_Computer.Hardware[1].Sensors[_VramFreeIndex].Value;
            }
            if (_GPUClockIndex == -1)
            {
                _AllCpu_GPU_Data_Safe[7] = -1;
            }
            else
            {
                _AllCpu_GPU_Data_Safe[7] = (float)_Computer.Hardware[1].Sensors[_GPUClockIndex].Value;
            }
            if (_GPUTemperatureIndex == -1)
            {
                _AllCpu_GPU_Data_Safe[8] = -1;
            }
            else
            {
                _AllCpu_GPU_Data_Safe[8] = (float)_Computer.Hardware[1].Sensors[_GPUTemperatureIndex].Value;
            }
        }
        return _AllCpu_GPU_Data_Safe;
    }

    public float[] GetAllCpu_GPU_Data()
    {
        _Computer.Hardware[0].Update();

        _AllCpu_GPU_Data[0] = GetMaxCPUClock();
        _AllCpu_GPU_Data[1] = (float)_Computer.Hardware[0].Sensors[_CPUTemperatureIndex].Value;
        _AllCpu_GPU_Data[2] = (float)_Computer.Hardware[0].Sensors[_CPUPowerIndex].Value;
        _AllCpu_GPU_Data[3] = (float)_Computer.Hardware[0].Sensors[_CPUUtilizationIndex].Value;
        _Computer.Hardware[1].Update();
        _AllCpu_GPU_Data[4] = (float)_Computer.Hardware[1].Sensors[_GPUPowerIndex].Value;
        _AllCpu_GPU_Data[5] = (float)_Computer.Hardware[1].Sensors[_GPUUtilizationIndex].Value;
        // _AllCpu_GPU_Data[6] = (float)_Computer.Hardware[1].Sensors[_VramUsedIndex].Value;
        _AllCpu_GPU_Data[6] = (float)_Computer.Hardware[1].Sensors[_VramFreeIndex].Value;
        _AllCpu_GPU_Data[7] = (float)_Computer.Hardware[1].Sensors[_GPUClockIndex].Value;
        _AllCpu_GPU_Data[8] = (float)_Computer.Hardware[1].Sensors[_GPUTemperatureIndex].Value;
        // if (_iGPUPowerIndex != -1)
        // {
        //     _AllCpu_IGPU_Data[4] = (float)_Computer.Hardware[0].Sensors[_iGPUPowerIndex].Value;
        // }


        return _AllCpu_GPU_Data;
    }
    // public int AddNullSensor(int Index) {
    //     int NewLength = _Computer.Hardware[Index].Sensors.Length+1;
    //     // Array.Resize(ref _Computer.Hardware[Index].Sensors, NewLength);
    //     _Computer.Hardware[Index].Sensors[NewLength].Value = -1;
    //     return NewLength;
    // }




    // private float[] _AlliGpuData;
    // public int InitGetAlliGpuData()
    // {
    //     _AlliGpuData = new float[2];
    //     D3DDisplayDevice.GetDeviceInfoByIdentifier(_D3DDisplayDeviceIdentifier, out D3DDisplayDevice.D3DDeviceInfo _deviceInfo);
    //     D3DDisplayDevice.D3DDeviceNodeInfo node = _deviceInfo.Nodes[0];
    //     _gpuNodeUsagePrevValue = node.RunningTime;
    //     _gpuNodeUsagePrevTick = node.QueryTime;
    //     return 1;
    // }
    // public float[] GetAlliGpuData()
    // {
    //     D3DDisplayDevice.GetDeviceInfoByIdentifier(_D3DDisplayDeviceIdentifier, out D3DDisplayDevice.D3DDeviceInfo _deviceInfo);
    //     _AlliGpuData[0] = _deviceInfo.GpuSharedUsed;

    //     D3DDisplayDevice.D3DDeviceNodeInfo node = _deviceInfo.Nodes[0];
    //     long runningTimeDiff = node.RunningTime - _gpuNodeUsagePrevValue;
    //     long timeDiff = node.QueryTime.Ticks - _gpuNodeUsagePrevTick.Ticks;
    //     float gpuNodeUsage = 100f * runningTimeDiff / timeDiff;
    //     _gpuNodeUsagePrevValue = node.RunningTime;
    //     _gpuNodeUsagePrevTick = node.QueryTime;
    //     _AlliGpuData[1] = gpuNodeUsage;
    //     return _AlliGpuData;
    // }





    // public float[] GetAllGpuData()
    // {
    //     _Computer.Hardware[1].Update();
    //     _AllGpuData[0] = (float)_Computer.Hardware[1].Sensors[_GPUClockIndex].Value;
    //     _AllGpuData[1] = (float)_Computer.Hardware[1].Sensors[_GPUTemperatureIndex].Value;
    //     _AllGpuData[2] = (float)_Computer.Hardware[1].Sensors[_GPUPowerIndex].Value;
    //     _AllGpuData[3] = (float)_Computer.Hardware[1].Sensors[_VramUsedIndex].Value;
    //     _AllGpuData[4] = (float)_Computer.Hardware[1].Sensors[_GPUUtilizationIndex].Value;

    //     return _AllGpuData;
    // }

    // public float[] GetAlliGpuData()
    // {
    //     _Computer.Hardware[1].Update();
    //     _AlliGpuData[0] = (float)_Computer.Hardware[1].Sensors[_GPUPowerIndex].Value;
    //     _AlliGpuData[1] = (float)_Computer.Hardware[1].Sensors[_VramUsedIndex].Value;
    //     _AlliGpuData[2] = (float)_Computer.Hardware[1].Sensors[_GPUUtilizationIndex].Value;

    //     return _AlliGpuData;
    // }

    public string textahk()
    {
        using StringWriter w = new(CultureInfo.InvariantCulture);
        w.WriteLine(_Computer.Hardware[1].HardwareType);
        return w.ToString();
    }
    public string GetGPUName()
    {
        using StringWriter w = new(CultureInfo.InvariantCulture);
        _GPUNameIndex = GetGPUNameIndex();
        if (_GPUNameIndex == -1)
        {
            w.WriteLine("-1");
        }
        else
        {
            w.WriteLine(_Computer.Hardware[_GPUNameIndex].HardwareType);
        }
        return w.ToString();
    }

    public int GetGPUNameIndex()
    {
        int NameIndex = 0;
        foreach (IGroup group in _groups)
        {
            foreach (IHardware hardware in group.Hardware)
            {
                //w.WriteLine(hardware.Identifier);
                if ((hardware.HardwareType == HardwareType.GpuNvidia) || (hardware.HardwareType == HardwareType.GpuIntel) || (hardware.HardwareType == HardwareType.GpuAmd))
                {
                    return NameIndex;
                }
                NameIndex++;
            }
        }
        return -1;
    }

    // public int GetiGPUPowerIndex()
    // {
    //     foreach (IGroup group in _groups)
    //     {
    //         foreach (IHardware hardware in group.Hardware)
    //         {
    //             //w.WriteLine(hardware.Identifier);
    //             if (hardware.HardwareType == HardwareType.Cpu)
    //             {
    //                 for (int j = 0; j < hardware.Sensors.Length; j++)
    //                 {
    //                     //找到温度传感器

    //                     if (hardware.Sensors[j].Name == "CPU Graphics")
    //                     {
    //                         if (hardware.Sensors[j].SensorType == SensorType.Power)
    //                         {
    //                             _iGPUPowerIndex = j;
    //                             return j;
    //                         }

    //                     }
    //                 }
    //             }
    //         }
    //     }
    //     return -1;
    // }






    public float GetCPUPower(IComputer Computer, int CPUPowerIndex)
    {
        return (float)Computer.Hardware[0].Sensors[CPUPowerIndex].Value;
    }
    public int GetCPUPowerIndex()
    {
        foreach (IGroup group in _groups)
        {
            foreach (IHardware hardware in group.Hardware)
            {
                //w.WriteLine(hardware.Identifier);
                if (hardware.HardwareType == HardwareType.Cpu)
                {
                    for (int j = 0; j < hardware.Sensors.Length; j++)
                    {
                        //找到温度传感器

                        if (hardware.Sensors[j].Name.Contains("Package"))
                        {
                            if (hardware.Sensors[j].SensorType == SensorType.Power)
                            {
                                return j;
                            }

                        }
                    }
                }
            }
        }
        return -1;
    }

    public int GetCPUClockIndex()
    {
        int CPUClockMaxNumber = 0;
        _CPUClockMinIndex = -1;
        foreach (IGroup group in _groups)
        {
            foreach (IHardware hardware in group.Hardware)
            {
                //w.WriteLine(hardware.Identifier);
                if (hardware.HardwareType == HardwareType.Cpu)
                {
                    for (int j = 0; j < hardware.Sensors.Length; j++)
                    {
                        if (hardware.Sensors[j].SensorType == SensorType.Clock)
                        {
                            if (hardware.Sensors[j].Name.Contains("Core"))
                            {
                                if (_CPUClockMinIndex == -1)
                                {
                                    _CPUClockMinIndex = j;
                                }
                                CPUClockMaxNumber = CPUClockMaxNumber + 1;
                            }
                        }
                    }
                    if (CPUClockMaxNumber != 0)
                    {
                        return CPUClockMaxNumber;
                    }
                }
            }
        }
        return -1;
    }
    public int UpdateCPUData(IComputer Computer)
    {
        Computer.Hardware[0].Update();
        return 1;
    }
    public float GetCPUTemperature(IComputer Computer, int CPUTemperatureIndex)
    {
        return (float)Computer.Hardware[0].Sensors[CPUTemperatureIndex].Value;
    }
    public int GetCPUTemperatureIndex()
    {
        int CPUTemperatureIndex_Safe = -1;
        foreach (IGroup group in _groups)
        {
            foreach (IHardware hardware in group.Hardware)
            {
                //w.WriteLine(hardware.Identifier);
                if (hardware.HardwareType == HardwareType.Cpu)
                {
                    for (int j = 0; j < hardware.Sensors.Length; j++)
                    {
                        //找到温度传感器
                        if (hardware.Sensors[j].SensorType == SensorType.Temperature)
                        {
                            if ((hardware.Sensors[j].Name == "CPU Package") || (hardware.Sensors[j].Name == "Core (Tctl/Tdie)"))
                            {
                                return j;
                            }

                            // if (hardware.Sensors[j].Name.Contains("Core"))
                            // {
                            //     CPUTemperatureIndex_Safe=j;
                            // }
                        }
                    }
                }
            }
        }
        return CPUTemperatureIndex_Safe;
    }


    public float GetCPUUtilization(IComputer Computer, int GetCPUUtilizationIndex)
    {
        return (float)Computer.Hardware[0].Sensors[GetCPUUtilizationIndex].Value;
    }
    public int GetCPUUtilizationIndex()
    {
        foreach (IGroup group in _groups)
        {
            foreach (IHardware hardware in group.Hardware)
            {
                //w.WriteLine(hardware.Identifier);
                if (hardware.HardwareType == HardwareType.Cpu)
                {
                    for (int j = 0; j < hardware.Sensors.Length; j++)
                    {
                        //找到温度传感器

                        if (hardware.Sensors[j].Name == "CPU Total")
                        {
                            if (hardware.Sensors[j].SensorType == SensorType.Load)
                            {
                                return j;
                            }

                        }
                    }
                }
            }
        }
        return -1;
    }


    public float GetGPUPower(IComputer Computer, int GPUPowerIndex)
    {
        return (float)Computer.Hardware[1].Sensors[GPUPowerIndex].Value;
    }
    public int GetGPUPowerIndex()
    {
        foreach (IGroup group in _groups)
        {
            foreach (IHardware hardware in group.Hardware)
            {
                //w.WriteLine(hardware.Identifier);
                if ((hardware.HardwareType == HardwareType.GpuNvidia) || (hardware.HardwareType == HardwareType.GpuIntel) || (hardware.HardwareType == HardwareType.GpuAmd))
                {
                    for (int j = 0; j < hardware.Sensors.Length; j++)
                    {
                        if (hardware.Sensors[j].SensorType == SensorType.Power)
                        {
                            if ((hardware.Sensors[j].Name == "GPU Package") || (hardware.Sensors[j].Name == "GPU Power"))
                            {
                                return j;
                            }

                        }
                    }
                }
            }
        }
        return -1;
    }
    public float GetGPUClock(IComputer Computer, int GPUClockIndex)
    {

        return (float)Computer.Hardware[1].Sensors[GPUClockIndex].Value;
    }
    public int GetGPUClockIndex()
    {
        foreach (IGroup group in _groups)
        {
            foreach (IHardware hardware in group.Hardware)
            {
                //w.WriteLine(hardware.Identifier);
                if ((hardware.HardwareType == HardwareType.GpuNvidia) || (hardware.HardwareType == HardwareType.GpuIntel) || (hardware.HardwareType == HardwareType.GpuAmd))
                {
                    for (int j = 0; j < hardware.Sensors.Length; j++)
                    {
                        if (hardware.Sensors[j].SensorType == SensorType.Clock)
                        {
                            if (hardware.Sensors[j].Name == "GPU Core")
                            {
                                return j;
                            }
                        }
                    }
                }
            }
        }
        return -1;
    }
    public int UpdateGPUData(IComputer Computer)
    {
        Computer.Hardware[1].Update();
        return 1;
    }
    public float GetGPUTemperature(IComputer Computer, int GPUTemperatureIndex)
    {
        return (float)Computer.Hardware[1].Sensors[GPUTemperatureIndex].Value;
    }
    public int GetGPUTemperatureIndex()
    {
        foreach (IGroup group in _groups)
        {
            foreach (IHardware hardware in group.Hardware)
            {
                //w.WriteLine(hardware.Identifier);
                if ((hardware.HardwareType == HardwareType.GpuNvidia) || (hardware.HardwareType == HardwareType.GpuIntel) || (hardware.HardwareType == HardwareType.GpuAmd))
                {
                    for (int j = 0; j < hardware.Sensors.Length; j++)
                    {
                        //找到温度传感器

                        if (hardware.Sensors[j].SensorType == SensorType.Temperature)
                        {
                            if (hardware.Sensors[j].Name == "GPU Hot Spot")
                            {
                                return j;
                            }

                        }
                    }
                }
            }
        }
        return -1;
    }

    public float GetGPUUtilization(IComputer Computer, int GetGPUUtilizationIndex)
    {
        return (float)Computer.Hardware[1].Sensors[GetGPUUtilizationIndex].Value;
    }
    public int GetGPUUtilizationIndex()
    {
        // using StringWriter w = new(CultureInfo.InvariantCulture);
        foreach (IGroup group in _groups)
        {
            foreach (IHardware hardware in group.Hardware)
            {
                if ((hardware.HardwareType == HardwareType.GpuNvidia) || (hardware.HardwareType == HardwareType.GpuIntel) || (hardware.HardwareType == HardwareType.GpuAmd))
                    for (int j = 0; j < hardware.Sensors.Length; j++)
                    {
                        if (hardware.Sensors[j].SensorType == SensorType.Load)
                        {
                            if ((hardware.Sensors[j].Name == "GPU Core") || (hardware.Sensors[j].Name == "D3D 3D") && (hardware.HardwareType == HardwareType.GpuIntel))
                            {
                                return j;
                            }
                        }
                    }
            }
        }
        return -1;
    }


    public float GetVramUtilization(IComputer Computer, int GetVramUsedIndex)
    {
        return (float)Computer.Hardware[1].Sensors[GetVramUsedIndex].Value;
    }
    public int GetVramUsedIndex()
    {
        foreach (IGroup group in _groups)
        {
            foreach (IHardware hardware in group.Hardware)
            {
                //w.WriteLine(hardware.Identifier);
                if ((hardware.HardwareType == HardwareType.GpuNvidia) || (hardware.HardwareType == HardwareType.GpuIntel) || (hardware.HardwareType == HardwareType.GpuAmd))
                {
                    for (int j = 0; j < hardware.Sensors.Length; j++)
                    {
                        if (hardware.Sensors[j].SensorType == SensorType.SmallData)
                        {
                            if ((hardware.Sensors[j].Name == "GPU Memory Used") || (hardware.Sensors[j].Name == "D3D Shared Memory Used") && (hardware.HardwareType == HardwareType.GpuIntel))
                            {
                                _VramUsedIndex = j;
                                return j;
                            }
                        }
                    }
                }
            }
        }
        return -1;
    }

    public int GetVramFreeIndex()
    {
        foreach (IGroup group in _groups)
        {
            foreach (IHardware hardware in group.Hardware)
            {
                //w.WriteLine(hardware.Identifier);
                if ((hardware.HardwareType == HardwareType.GpuNvidia) || (hardware.HardwareType == HardwareType.GpuIntel) || (hardware.HardwareType == HardwareType.GpuAmd))
                {
                    for (int j = 0; j < hardware.Sensors.Length; j++)
                    {
                        if (hardware.Sensors[j].SensorType == SensorType.SmallData)
                        {
                            if ((hardware.Sensors[j].Name == "GPU Memory Free") || (hardware.Sensors[j].Name == "D3D Shared Memory Free"))
                            {
                                return j;
                            }
                        }
                    }
                }
            }
        }
        return -1;
    }
    public int GetVramTotalIndex()
    {
        foreach (IGroup group in _groups)
        {
            foreach (IHardware hardware in group.Hardware)
            {
                //w.WriteLine(hardware.Identifier);
                if ((hardware.HardwareType == HardwareType.GpuNvidia) || (hardware.HardwareType == HardwareType.GpuIntel) || (hardware.HardwareType == HardwareType.GpuAmd))
                {
                    for (int j = 0; j < hardware.Sensors.Length; j++)
                    {
                        //找到温度传感器

                        if (hardware.Sensors[j].Name == "GPU Memory Total")
                        {
                            if (hardware.Sensors[j].SensorType == SensorType.SmallData)
                            {
                                _VramTotalIndex = j;
                                return j;
                            }

                        }
                    }
                }
            }
        }
        return -1;
    }


    // public float GetiGpuUsage(string D3DDisplayDeviceIdentifier)
    // {
    //     //if (D3DDisplayDevice.GetDeviceInfoByIdentifier(D3DDisplayDeviceIdentifier, out D3DDisplayDevice.D3DDeviceInfo deviceInfo))
    //     //{
    //     //int i = 0;
    //     //deviceInfo.Nodes[1];
    //     //D3DDisplayDevice.GetDeviceInfoByIdentifier(D3DDisplayDeviceIdentifier, out D3DDisplayDevice.D3DDeviceInfo deviceInfo);
    //     //long gpuNodeUsagePrevValue;
    //     //DateTime gpuNodeUsagePrevTick;
    //     //foreach (D3DDisplayDevice.D3DDeviceNodeInfo node in deviceInfo.Nodes)
    //     //{

    //     D3DDisplayDevice.D3DDeviceNodeInfo node = _deviceInfo.Nodes[0];

    //     long runningTimeDiff = node.RunningTime - _gpuNodeUsagePrevValue;
    //     long timeDiff = node.QueryTime.Ticks - _gpuNodeUsagePrevTick.Ticks;

    //     float gpuNodeUsage = 100f * runningTimeDiff / timeDiff;
    //     _gpuNodeUsagePrevValue = node.RunningTime;
    //     _gpuNodeUsagePrevTick = node.QueryTime;
    //     // }

    //     return gpuNodeUsage;

    //     //}
    //     // int i = -1;
    //     //return -1;
    // }









    public float GetGpuSharedMemoryUsed(string D3DDisplayDeviceIdentifier)
    {
        if (D3DDisplayDevice.GetDeviceInfoByIdentifier(D3DDisplayDeviceIdentifier, out D3DDisplayDevice.D3DDeviceInfo deviceInfo))
        {
            float gpuSharedMemoryUsage = deviceInfo.GpuSharedUsed;
            return gpuSharedMemoryUsage;
        }
        return 0;
    }

    public float GetGpuSharedLimit(string D3DDisplayDeviceIdentifier)
    {
        if (D3DDisplayDevice.GetDeviceInfoByIdentifier(D3DDisplayDeviceIdentifier, out D3DDisplayDevice.D3DDeviceInfo deviceInfo))
        {
            float gpuSharedMemoryUsage = deviceInfo.GpuSharedLimit;
            return gpuSharedMemoryUsage;
        }
        return 0;
    }

    public string GetD3DDisplayDeviceIdentifier(int index)
    {
        _D3DDisplayDeviceIdentifier = D3DDisplayDevice.GetDeviceIdentifiers()[index];
        return _D3DDisplayDeviceIdentifier;
    }


    public string[] GetSmartData()
    {
        List<string> SmartDataList = new List<string>();
        foreach (IGroup group in _groups)
        {
            foreach (IHardware hardware in group.Hardware)
            {
                if (hardware.HardwareType == HardwareType.Storage)
                {
                    SmartDataList.Add(hardware.GetReport());
                }
            }
        }
        string[] SmartData = new string[SmartDataList.Count];
        for (int i = 0; i < SmartDataList.Count; i++)
        {
            SmartData[i] = SmartDataList[i];
        }

        return SmartData;
    }

    public string GetMotherboardInfo()
    {
        foreach (IGroup group in _groups)
        {
            foreach (IHardware hardware in group.Hardware)
            {
                if (hardware.HardwareType == HardwareType.Motherboard)
                {
                    return hardware.GetReport();
                }
            }
        }
        return "-1";
    }

    public string GetCpuInfo()
    {
        return _Computer.Hardware[0].GetReport();
    }

    public string GetCPUIDInfo()
    {
        return _groups[0].GetReport();
    }

    public string GetGPUInfo()
    {
        return _Computer.Hardware[1].GetReport();
    }

    public string GetD3DInfo()
    {
        foreach (IGroup group in _groups)
        {
            String Report = group.GetReport();
            if (Report.Contains("Intel GPU (D3D)"))
            {
                return Report;
            }
        }
        return "-1";
    }

}

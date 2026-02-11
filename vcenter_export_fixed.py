#!/usr/bin/env python3
"""
Script untuk mengekspor semua data vCenter ke CSV
Version: 1.1 (Fixed - Handle deleted/incomplete objects)
"""

from pyVim.connect import SmartConnect, Disconnect
from pyVmomi import vim
import ssl
import csv
import atexit
import urllib3
from datetime import datetime

# Disable SSL warnings
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# ============================================
# KONFIGURASI - EDIT BAGIAN INI
# ============================================
VCENTER_HOST = "xx.xx.xx.xx"
VCENTER_USER = "inisial"
VCENTER_PASSWORD = "password"
VCENTER_PORT = 443

# ============================================
# FUNGSI HELPER
# ============================================

def connect_vcenter():
    """Koneksi ke vCenter"""
    print(f"Menghubungkan ke vCenter: {VCENTER_HOST}...")
    context = ssl._create_unverified_context()
    si = SmartConnect(
        host=VCENTER_HOST,
        user=VCENTER_USER,
        pwd=VCENTER_PASSWORD,
        port=VCENTER_PORT,
        sslContext=context
    )
    atexit.register(Disconnect, si)
    print("✓ Koneksi berhasil!\n")
    return si

def get_all_objs(content, vimtype):
    """Mengambil semua objek dengan tipe tertentu - dengan error handling"""
    obj = {}
    container = content.viewManager.CreateContainerView(
        content.rootFolder, vimtype, True
    )
    for managed_object_ref in container.view:
        try:
            # Coba akses name untuk validasi objek masih ada
            name = managed_object_ref.name
            obj.update({managed_object_ref: name})
        except Exception as e:
            # Skip objek yang sudah dihapus atau belum selesai dibuat
            if "ManagedObjectNotFound" in str(e) or "has already been deleted" in str(e):
                print(f"  ⚠ Skipping deleted/incomplete object")
            else:
                print(f"  ⚠ Skipping object due to: {e}")
            continue
    container.Destroy()
    return obj

def write_csv(filename, data, fieldnames=None):
    """Menulis data ke CSV"""
    if not data:
        print(f"  ⚠ Tidak ada data untuk {filename}")
        return

    if fieldnames is None:
        fieldnames = data[0].keys()

    with open(filename, 'w', newline='', encoding='utf-8') as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(data)
    print(f"  ✓ {filename} ({len(data)} records)")

def safe_get_property(obj, property_chain, default='N/A'):
    """Safely get nested property with fallback"""
    try:
        result = obj
        for prop in property_chain.split('.'):
            result = getattr(result, prop)
        return result if result is not None else default
    except:
        return default

# ============================================
# FUNGSI EKSPOR DATA
# ============================================

def export_vcenter_info(content):
    """1. Export vCenter Info"""
    print("Mengekspor vCenter Info...")
    try:
        data = [{
            'name': content.about.name,
            'version': content.about.version,
            'build': content.about.build,
            'os_type': content.about.osType,
            'api_type': content.about.apiType,
            'instance_uuid': content.about.instanceUuid
        }]
        write_csv('vInfo.csv', data)
    except Exception as e:
        print(f"  ✗ Error: {e}")

def export_clusters(content):
    """2. Export Clusters"""
    print("Mengekspor Clusters...")
    data = []
    for cluster in get_all_objs(content, [vim.ClusterComputeResource]):
        try:
            data.append({
                'name': cluster.name,
                'total_cpu_cores': safe_get_property(cluster, 'summary.numCpuCores', 0),
                'total_cpu_threads': safe_get_property(cluster, 'summary.numCpuThreads', 0),
                'total_memory_gb': round(safe_get_property(cluster, 'summary.totalMemory', 0) / 1024**3, 2),
                'num_hosts': safe_get_property(cluster, 'summary.numHosts', 0),
                'num_effective_hosts': safe_get_property(cluster, 'summary.numEffectiveHosts', 0),
                'drs_enabled': safe_get_property(cluster, 'configuration.drsConfig.enabled', False),
                'drs_behavior': safe_get_property(cluster, 'configuration.drsConfig.defaultVmBehavior', 'N/A'),
                'ha_enabled': safe_get_property(cluster, 'configuration.dasConfig.enabled', False),
                'overall_status': safe_get_property(cluster, 'overallStatus', 'N/A')
            })
        except Exception as e:
            print(f"  ⚠ Error pada cluster {cluster.name}: {e}")
    write_csv('vCluster.csv', data)

def export_hosts(content):
    """3. Export Hosts"""
    print("Mengekspor Hosts...")
    data = []
    for host in get_all_objs(content, [vim.HostSystem]):
        try:
            cpu_pkg = host.hardware.cpuPkg[0].description if host.hardware.cpuPkg else 'N/A'
            data.append({
                'name': host.name,
                'manufacturer': safe_get_property(host, 'hardware.systemInfo.vendor', 'N/A'),
                'model': safe_get_property(host, 'hardware.systemInfo.model', 'N/A'),
                'cpu_model': cpu_pkg,
                'cpu_cores': safe_get_property(host, 'hardware.cpuInfo.numCpuCores', 0),
                'cpu_threads': safe_get_property(host, 'hardware.cpuInfo.numCpuThreads', 0),
                'cpu_mhz': safe_get_property(host, 'hardware.cpuInfo.hz', 0) // 1000000,
                'memory_gb': round(safe_get_property(host, 'hardware.memorySize', 0) / 1024**3, 2),
                'num_nics': len(host.config.network.pnic) if host.config.network.pnic else 0,
                'connection_state': safe_get_property(host, 'runtime.connectionState', 'N/A'),
                'power_state': safe_get_property(host, 'runtime.powerState', 'N/A'),
                'maintenance_mode': safe_get_property(host, 'runtime.inMaintenanceMode', False),
                'version': safe_get_property(host, 'config.product.version', 'N/A'),
                'build': safe_get_property(host, 'config.product.build', 'N/A'),
                'overall_status': safe_get_property(host, 'overallStatus', 'N/A')
            })
        except Exception as e:
            print(f"  ⚠ Error pada host: {e}")
    write_csv('vHost.csv', data)

def export_datastores(content):
    """4. Export Datastores"""
    print("Mengekspor Datastores...")
    data = []
    for ds in get_all_objs(content, [vim.Datastore]):
        try:
            capacity_gb = round(safe_get_property(ds, 'summary.capacity', 0) / 1024**3, 2)
            free_gb = round(safe_get_property(ds, 'summary.freeSpace', 0) / 1024**3, 2)
            used_gb = capacity_gb - free_gb
            used_percent = round((used_gb / capacity_gb * 100), 2) if capacity_gb > 0 else 0

            data.append({
                'name': ds.name,
                'type': safe_get_property(ds, 'summary.type', 'N/A'),
                'capacity_gb': capacity_gb,
                'free_gb': free_gb,
                'used_gb': used_gb,
                'used_percent': used_percent,
                'accessible': safe_get_property(ds, 'summary.accessible', False),
                'multiple_host_access': safe_get_property(ds, 'summary.multipleHostAccess', False),
                'maintenance_mode': safe_get_property(ds, 'summary.maintenanceMode', 'N/A'),
                'uncommitted_gb': round(safe_get_property(ds, 'summary.uncommitted', 0) / 1024**3, 2),
                'num_vms': len(ds.vm) if ds.vm else 0
            })
        except Exception as e:
            print(f"  ⚠ Error pada datastore: {e}")
    write_csv('vDatastore.csv', data)

def export_vms(content):
    """5. Export VMs"""
    print("Mengekspor VMs...")
    data = []
    skipped = 0

    for vm in get_all_objs(content, [vim.VirtualMachine]):
        try:
            # Skip templates
            if safe_get_property(vm, 'config.template', False):
                continue

            # Count disks and NICs safely
            num_disks = 0
            num_nics = 0
            try:
                if vm.config and vm.config.hardware and vm.config.hardware.device:
                    for device in vm.config.hardware.device:
                        if isinstance(device, vim.vm.device.VirtualDisk):
                            num_disks += 1
                        elif isinstance(device, vim.vm.device.VirtualEthernetCard):
                            num_nics += 1
            except:
                pass

            data.append({
                'name': vm.name,
                'power_state': safe_get_property(vm, 'runtime.powerState', 'N/A'),
                'num_cpu': safe_get_property(vm, 'config.hardware.numCPU', 0),
                'num_cores_per_socket': safe_get_property(vm, 'config.hardware.numCoresPerSocket', 0),
                'memory_mb': safe_get_property(vm, 'config.hardware.memoryMB', 0),
                'memory_gb': round(safe_get_property(vm, 'config.hardware.memoryMB', 0) / 1024, 2),
                'guest_os': safe_get_property(vm, 'config.guestFullName', 'N/A'),
                'guest_os_id': safe_get_property(vm, 'config.guestId', 'N/A'),
                'version': safe_get_property(vm, 'config.version', 'N/A'),
                'tools_status': safe_get_property(vm, 'guest.toolsStatus', 'N/A'),
                'tools_version': safe_get_property(vm, 'guest.toolsVersion', 'N/A'),
                'host': vm.runtime.host.name if vm.runtime and vm.runtime.host else 'N/A',
                'num_disks': num_disks,
                'num_nics': num_nics,
                'overall_status': safe_get_property(vm, 'overallStatus', 'N/A'),
                'annotation': safe_get_property(vm, 'config.annotation', '')
            })
        except Exception as e:
            skipped += 1
            print(f"  ⚠ Skipping VM due to error: {str(e)[:80]}")

    write_csv('vVM.csv', data)
    if skipped > 0:
        print(f"  ℹ Skipped {skipped} VMs due to errors")

def export_disks(content):
    """6. Export Virtual Disks"""
    print("Mengekspor Virtual Disks...")
    data = []
    for vm in get_all_objs(content, [vim.VirtualMachine]):
        try:
            if safe_get_property(vm, 'config.template', False):
                continue

            if not vm.config or not vm.config.hardware or not vm.config.hardware.device:
                continue

            for device in vm.config.hardware.device:
                try:
                    if isinstance(device, vim.vm.device.VirtualDisk):
                        datastore_name = 'N/A'
                        if hasattr(device.backing, 'datastore') and device.backing.datastore:
                            try:
                                datastore_name = device.backing.datastore.name
                            except:
                                pass

                        data.append({
                            'vm_name': vm.name,
                            'label': safe_get_property(device, 'deviceInfo.label', 'N/A'),
                            'capacity_gb': round(safe_get_property(device, 'capacityInKB', 0) / 1024**2, 2),
                            'capacity_mb': round(safe_get_property(device, 'capacityInKB', 0) / 1024, 2),
                            'disk_mode': safe_get_property(device, 'backing.diskMode', 'N/A'),
                            'thin_provisioned': safe_get_property(device, 'backing.thinProvisioned', False),
                            'disk_type': type(device.backing).__name__ if device.backing else 'N/A',
                            'datastore': datastore_name,
                            'controller': safe_get_property(device, 'controllerKey', 'N/A'),
                            'unit_number': safe_get_property(device, 'unitNumber', 'N/A')
                        })
                except:
                    continue
        except Exception as e:
            print(f"  ⚠ Error processing disks for VM: {str(e)[:60]}")
    write_csv('vDisk.csv', data)

def export_snapshots(content):
    """7. Export Snapshots"""
    print("Mengekspor Snapshots...")
    data = []

    def process_snapshot(vm_name, snapshot_tree, parent_name=''):
        """Recursive function untuk memproses snapshot tree"""
        for snapshot in snapshot_tree:
            try:
                create_time = snapshot.createTime.strftime('%Y-%m-%d %H:%M:%S') if snapshot.createTime else 'N/A'
                data.append({
                    'vm_name': vm_name,
                    'snapshot_name': snapshot.name,
                    'description': snapshot.description if snapshot.description else '',
                    'create_time': create_time,
                    'state': snapshot.state,
                    'quiesced': snapshot.quiesced,
                    'parent_snapshot': parent_name,
                    'id': snapshot.id
                })
                # Process child snapshots
                if snapshot.childSnapshotList:
                    process_snapshot(vm_name, snapshot.childSnapshotList, snapshot.name)
            except:
                continue

    for vm in get_all_objs(content, [vim.VirtualMachine]):
        try:
            if vm.snapshot:
                process_snapshot(vm.name, vm.snapshot.rootSnapshotList)
        except Exception as e:
            pass

    write_csv('vSnapshot.csv', data)

def export_standard_portgroups(content):
    """8. Export Standard Port Groups"""
    print("Mengekspor Standard Port Groups...")
    data = []
    for host in get_all_objs(content, [vim.HostSystem]):
        try:
            if not host.config or not host.config.network or not host.config.network.portgroup:
                continue

            for pg in host.config.network.portgroup:
                try:
                    num_active_nics = 0
                    try:
                        if pg.spec.policy.nicTeaming.nicOrder.activeNic:
                            num_active_nics = len(pg.spec.policy.nicTeaming.nicOrder.activeNic)
                    except:
                        pass

                    data.append({
                        'host': host.name,
                        'name': pg.spec.name,
                        'vlan_id': pg.spec.vlanId,
                        'vswitch': pg.spec.vswitchName,
                        'num_ports': num_active_nics,
                        'security_allow_promiscuous': safe_get_property(pg, 'spec.policy.security.allowPromiscuous', False),
                        'security_mac_changes': safe_get_property(pg, 'spec.policy.security.macChanges', False),
                        'security_forged_transmits': safe_get_property(pg, 'spec.policy.security.forgedTransmits', False)
                    })
                except:
                    continue
        except Exception as e:
            print(f"  ⚠ Error processing portgroups for host: {str(e)[:60]}")
    write_csv('vPortgroup_Std.csv', data)

def export_distributed_portgroups(content):
    """9. Export Distributed Port Groups"""
    print("Mengekspor Distributed Port Groups...")
    data = []
    for dvpg in get_all_objs(content, [vim.dvs.DistributedVirtualPortgroup]):
        try:
            config = dvpg.config

            # Get VLAN info
            vlan_id = 'N/A'
            vlan_type = 'N/A'
            try:
                if hasattr(config.defaultPortConfig, 'vlan'):
                    vlan_config = config.defaultPortConfig.vlan
                    if isinstance(vlan_config, vim.dvs.VmwareDistributedVirtualSwitch.VlanIdSpec):
                        vlan_id = vlan_config.vlanId
                        vlan_type = 'VLAN'
                    elif isinstance(vlan_config, vim.dvs.VmwareDistributedVirtualSwitch.TrunkVlanSpec):
                        vlan_id = str([f"{r.start}-{r.end}" for r in vlan_config.vlanId])
                        vlan_type = 'Trunk'
            except:
                pass

            dvs_name = 'N/A'
            try:
                dvs_name = config.distributedVirtualSwitch.name
            except:
                pass

            data.append({
                'name': dvpg.name,
                'dvswitch': dvs_name,
                'type': safe_get_property(config, 'type', 'N/A'),
                'num_ports': safe_get_property(config, 'numPorts', 0),
                'vlan_id': vlan_id,
                'vlan_type': vlan_type,
                'port_binding': safe_get_property(config, 'defaultPortConfig.portBindingType', 'N/A'),
                'auto_expand': safe_get_property(config, 'autoExpand', False)
            })
        except Exception as e:
            print(f"  ⚠ Error on distributed portgroup: {str(e)[:60]}")
    write_csv('vPortgroup_DV.csv', data)

def export_standard_vswitches(content):
    """10. Export Standard vSwitches"""
    print("Mengekspor Standard vSwitches...")
    data = []
    for host in get_all_objs(content, [vim.HostSystem]):
        try:
            if not host.config or not host.config.network or not host.config.network.vswitch:
                continue

            for vsw in host.config.network.vswitch:
                try:
                    pnic_list = ','.join(vsw.pnic) if vsw.pnic else ''
                    data.append({
                        'host': host.name,
                        'name': vsw.name,
                        'num_ports': vsw.spec.numPorts,
                        'num_ports_available': vsw.numPortsAvailable,
                        'mtu': vsw.mtu,
                        'num_physical_nics': len(vsw.pnic) if vsw.pnic else 0,
                        'physical_nics': pnic_list,
                        'num_portgroups': len(vsw.portgroup) if vsw.portgroup else 0
                    })
                except:
                    continue
        except Exception as e:
            print(f"  ⚠ Error processing vswitches for host: {str(e)[:60]}")
    write_csv('vSwitch_Std.csv', data)

def export_vmkernel_nics(content):
    """11. Export VMkernel NICs"""
    print("Mengekspor VMkernel NICs...")
    data = []
    for host in get_all_objs(content, [vim.HostSystem]):
        try:
            if not host.config or not host.config.network or not host.config.network.vnic:
                continue

            for vnic in host.config.network.vnic:
                try:
                    dvport_id = 'N/A'
                    try:
                        if hasattr(vnic.spec, 'distributedVirtualPort') and vnic.spec.distributedVirtualPort:
                            dvport_id = vnic.spec.distributedVirtualPort.portKey
                    except:
                        pass

                    data.append({
                        'host': host.name,
                        'device': vnic.device,
                        'portgroup': vnic.portgroup,
                        'dvport_id': dvport_id,
                        'mac': safe_get_property(vnic, 'spec.mac', 'N/A'),
                        'ip': safe_get_property(vnic, 'spec.ip.ipAddress', 'N/A'),
                        'subnet_mask': safe_get_property(vnic, 'spec.ip.subnetMask', 'N/A'),
                        'dhcp': safe_get_property(vnic, 'spec.ip.dhcp', False),
                        'mtu': safe_get_property(vnic, 'spec.mtu', 1500)
                    })
                except:
                    continue
        except Exception as e:
            print(f"  ⚠ Error processing vmkernel NICs for host: {str(e)[:60]}")
    write_csv('vVMkernelNIC.csv', data)

def export_physical_nics(content):
    """12. Export Physical NICs"""
    print("Mengekspor Physical NICs...")
    data = []
    for host in get_all_objs(content, [vim.HostSystem]):
        try:
            if not host.config or not host.config.network or not host.config.network.pnic:
                continue

            for pnic in host.config.network.pnic:
                try:
                    speed = 'Down'
                    duplex = 'N/A'
                    if pnic.linkSpeed:
                        speed = pnic.linkSpeed.speedMb
                        duplex = pnic.linkSpeed.duplex

                    vswitch_name = 'Not assigned'
                    try:
                        if host.config.network.vswitch:
                            for vsw in host.config.network.vswitch:
                                if vsw.pnic and pnic.key in vsw.pnic:
                                    vswitch_name = vsw.name
                                    break
                    except:
                        pass

                    data.append({
                        'host': host.name,
                        'device': pnic.device,
                        'mac': pnic.mac,
                        'pci': pnic.pci,
                        'driver': pnic.driver,
                        'link_speed_mb': speed,
                        'duplex': duplex,
                        'wol_supported': pnic.wakeOnLanSupported,
                        'vswitch': vswitch_name
                    })
                except:
                    continue
        except Exception as e:
            print(f"  ⚠ Error processing physical NICs for host: {str(e)[:60]}")
    write_csv('vPNIC.csv', data)

def export_hbas(content):
    """13. Export HBAs"""
    print("Mengekspor HBAs...")
    data = []
    for host in get_all_objs(content, [vim.HostSystem]):
        try:
            if not host.config or not host.config.storageDevice or not host.config.storageDevice.hostBusAdapter:
                continue

            for hba in host.config.storageDevice.hostBusAdapter:
                try:
                    hba_type = 'Unknown'
                    wwn = 'N/A'
                    speed = 'N/A'

                    if isinstance(hba, vim.host.FibreChannelHba):
                        hba_type = 'Fibre Channel'
                        wwn = ':'.join([hba.portWorldWideName[i:i+2] for i in range(0, len(hba.portWorldWideName), 2)])
                        speed = hba.speed if hasattr(hba, 'speed') else 'N/A'
                    elif isinstance(hba, vim.host.InternetScsiHba):
                        hba_type = 'iSCSI'
                        wwn = hba.iScsiName if hasattr(hba, 'iScsiName') else 'N/A'
                    elif isinstance(hba, vim.host.ParallelScsiHba):
                        hba_type = 'Parallel SCSI'

                    data.append({
                        'host': host.name,
                        'device': hba.device,
                        'type': hba_type,
                        'model': hba.model,
                        'driver': hba.driver,
                        'pci': hba.pci,
                        'status': hba.status,
                        'wwn_or_iqn': wwn,
                        'speed': speed
                    })
                except:
                    continue
        except Exception as e:
            print(f"  ⚠ Error processing HBAs for host: {str(e)[:60]}")
    write_csv('vHBA.csv', data)

# ============================================
# MAIN FUNCTION
# ============================================

def main():
    """Fungsi utama"""
    print("="*60)
    print("vCenter Data Exporter ke CSV")
    print("="*60)
    print(f"Timestamp: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")

    try:
        # Koneksi ke vCenter
        si = connect_vcenter()
        content = si.RetrieveContent()

        # Export semua data
        print("Memulai ekspor data...\n")

        export_vcenter_info(content)
        export_clusters(content)
        export_hosts(content)
        export_datastores(content)
        export_vms(content)
        export_disks(content)
        export_snapshots(content)
        export_standard_portgroups(content)
        export_distributed_portgroups(content)
        export_standard_vswitches(content)
        export_vmkernel_nics(content)
        export_physical_nics(content)
        export_hbas(content)

        print("\n" + "="*60)
        print("✓ SEMUA DATA BERHASIL DIEKSPOR!")
        print("="*60)
        print("\nFile yang dihasilkan:")
        print("  1. vInfo.csv")
        print("  2. vCluster.csv")
        print("  3. vHost.csv")
        print("  4. vDatastore.csv")
        print("  5. vVM.csv")
        print("  6. vDisk.csv")
        print("  7. vSnapshot.csv")
        print("  8. vPortgroup_Std.csv")
        print("  9. vPortgroup_DV.csv")
        print(" 10. vSwitch_Std.csv")
        print(" 11. vVMkernelNIC.csv")
        print(" 12. vPNIC.csv")
        print(" 13. vHBA.csv")

    except Exception as e:
        print(f"\n✗ ERROR: {e}")
        import traceback
        traceback.print_exc()

    print("\nSelesai!")

if __name__ == "__main__":
    main()

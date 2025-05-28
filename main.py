
import os
import contextlib
import importlib
from ctypes import POINTER, cast
from comtypes import GUID, IUnknown, COMMETHOD, HRESULT, BSTR
from comtypes.automation import VARIANT
from comtypes.client import CreateObject
import comtypes.client

try:
    importlib.import_module('comtypes.gen.DirectShowLib')
except ImportError:
    @contextlib.contextmanager
    def _working_directory(dir):
        prev_cwd = os.getcwd()
        os.chdir(dir)
        try:
            yield
        finally:
            os.chdir(prev_cwd)

    # generate the DirectShowLib module if it doesn't exist
    with _working_directory(os.path.dirname(os.path.abspath(__file__))):
        comtypes.client.GetModule('DirectShow.tlb')

from comtypes.gen.DirectShowLib import (
    FilterGraph,
    CaptureGraphBuilder2,
    ICreateDevEnum,
    IBaseFilter,
    IBindCtx,
    IMoniker,
    IAMCameraControl
)

_quartz = comtypes.client.GetModule('quartz.dll')
IMediaControl = _quartz.IMediaControl

PIN_CATEGORY_CAPTURE = GUID('{fb6c4281-0353-11d1-905f-0000c0cc16ba}')
MEDIATYPE_Interleaved = GUID('{73766169-0000-0010-8000-00aa00389b71}')
MEDIATYPE_Video = GUID('{73646976-0000-0010-8000-00AA00389B71}')

CLSID_VideoInputDeviceCategory = GUID("{860BB310-5D01-11d0-BD3B-00A0C911CE86}")
CLSID_SystemDeviceEnum = GUID('{62BE5D10-60EB-11d0-BD3B-00A0C911CE86}')
IID_IPropertyBag = GUID("{55272A00-42CB-11CE-8135-00AA004BB851}")
IID_ICreateDevEnum = GUID("{29840822-5B84-11D0-BD3B-00A0C911CE86}")
IID_IEnumMoniker = GUID("{00000102-0000-0000-C000-000000000046}")

CameraControl_Focus = 0x6
CameraControl_Flags_Auto = 0x1
CameraControl_Flags_Manual = 0x2

filter_graph = CreateObject(FilterGraph)
control = filter_graph.QueryInterface(IMediaControl)
graph_builder = CreateObject(CaptureGraphBuilder2)
graph_builder.SetFiltergraph(filter_graph)
dev_enum = CreateObject(CLSID_SystemDeviceEnum, interface=ICreateDevEnum)
class_enum = dev_enum.CreateClassEnumerator(CLSID_VideoInputDeviceCategory, 0)

class IPropertyBag(IUnknown):
    _iid_ = IID_IPropertyBag
    _methods_ = [
        COMMETHOD([], HRESULT, "Read",
                  (["in"], BSTR, "propName"),
                  (["out"], POINTER(VARIANT), "pVar"),
                  (["in"], POINTER(IUnknown), "pErrorLog")),
    ]

monikers = []
(moniker, fetched) = class_enum.RemoteNext(1)
while fetched == 1:
    prop_bag = moniker.QueryInterface(IPropertyBag)
    name_var = prop_bag.Read("FriendlyName", None)
    monikers.append((moniker, name_var))
    print(f"Found camera: {name_var}")
    (moniker, fetched) = class_enum.RemoteNext(1)

print("Available cameras:")
for i, (moniker, name_var) in enumerate(monikers):
    print(f"    {i}: {name_var}")

camera_index = int(input("Select camera index: "))
moniker = monikers[camera_index][0]
print(f"Selected camera: {monikers[camera_index][1]}")

null_context = POINTER(IBindCtx)()
null_moniker = POINTER(IMoniker)()
source = moniker.RemoteBindToObject(
    null_context, null_moniker, IBaseFilter._iid_
)

camera_control: IAMCameraControl = cast(
    graph_builder.RemoteFindInterface(
        PIN_CATEGORY_CAPTURE,
        MEDIATYPE_Video,
        source,
        IAMCameraControl._iid_,
    ),
    POINTER(IAMCameraControl)
).value

min_f, max_f, step, default_f, caps = camera_control.GetRange(CameraControl_Focus)
f_val, f_flags = camera_control.Get(CameraControl_Focus)
print("Camera Focus:")
print(f"    Range: {min_f} - {max_f} with step size {step}")
print(f"    Default: {default_f}")
print(f"    Current value: {f_val}")
print(f"    Current mode: {'Auto' if f_flags & CameraControl_Flags_Auto else 'Manual'}")

new_focus = input("Input new focus value (or press Enter to skip): ")
if new_focus:
    new_focus = int(new_focus)
    if min_f <= new_focus <= max_f:
        new_focus = int(round((new_focus - min_f) / step) * step) + min_f
        print("Setting focus to:", new_focus)
        camera_control.Set(CameraControl_Focus, new_focus, CameraControl_Flags_Manual)
    else:
        print(f"Focus value must be between {min_f} and {max_f}")

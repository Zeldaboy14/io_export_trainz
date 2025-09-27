  # -*- coding: utf-8 -*-

__author__ = ["uschi0815"]
__url__ = ("https://sourceforge.net/projects/blenderextrainz/")
__version__ = "0.96"
__bpydoc__ = """\

Blender Exporter for Trainz

This script exports a Blender scene to use it in Trainz. It is based on
Blenders PythonAPI and the TrainzMeshImporter from N3V.

Special thanks goes to decapod for sharing his knowledge with me, johnwhelan
for boosting my english and this project to sourceforge and Peter V. for his
awesome "Mesh Text Viewer".

known bugs:
 - currently none

hints:
  In Trainz models you can have two kinds of special objects: attachment points
  and bones. For attachment points the exporter expects objects of type
  "Empty"; for bones objects of type "Lattice" and "Armature" will be accepted.
  All of them need to be named according the naming conventions in effect for
  attachements( a.* ) and bones( b.r.* ). Objects of other type and objects
  with proper type but names not matching this conventions will be ignored as
  attachment points or bones.

  If you want to export animations with events, than you have to create a
  Blender TX object and enter the lines as you would do to create an event
  file. By renaming this object to "events" you ensure the exporter will not
  ignore your event list.

  It's possible to make parent-child-relations between meshes to simplify the
  design process. However, the chain must be end with a valid parent bone.

legal:
  all trade marks are properties of the respective owners

"""

# ***** BEGIN GPL LICENSE BLOCK *****
#
# This program is free software; you can redistribute it and/or
# modify it under the terms of the GNU General Public License
# as published by the Free Software Foundation; either version 2
# of the License, or (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with this program; if not, write to the Free Software Foundation,
# Inc., 59 Temple Place - Suite 330, Boston, MA  02111-1307, USA.
#
# ***** END GPL LICENCE BLOCK *****
# --------------------------------------------------------------------------

# <pep8 compliant>

### todo
# - find a way to get rid of the assumption that we use 8 bit per plane (RGBA);
#   see use_alpha - check


### changes in 0.96
# - Blenders texture option Image:Use Alpha is now also utilized to trigger
#   the Alpha= line in texture.txt file
# - change property name TextureSlot.colorspec_factor(Blender 2.5x?) to
#   TextureSlot.specular_color_factor
# - scaling removed; With Blenders capability to display metrical or imperial
#   units scaling is only "nice to have" and not longer mandatory (should have
#   done this long before)
# - missleading error messages related to faulty materials and a bug which
#   throw away vertex group "ERROR_no_Material_assigned" short after it's
#   creation could be fixed - many thanks to mjolnir
# - now [&><] get transformed into character entity references before dropping
#   them into the XML file
# - added event trigger to the "names-check"
# - precision set to 4 -> 0.1 Millimeter precision should be sufficient for
#   any Trainz asset
# - widen the animation data check to find at least one object with key frames
# - add the 'conveniance' dot to exporter generated material decorations as the
#   length of material names is no longer limited to 31 characters 

### changes in 0.95
# - XML strings will be written directly to disk to reduce memory consumption;
#   so far I've not noticed any slow down
# - material opacity is set to 1 regardless of alpha slider value if
#   transparency is not enabled
# - changes needed by switching from faces to polygons implemented - many
#   thanks to pcas1986
# - logging of script options


bl_info = {
    "name": "TRAINZ Exporter",
    "description": "Export objects as XML/IM/KIN file to import into Trainz",
    "author": "uschi0815",
    "version": (0, 96),
    "blender": (2, 59, 4),
    "api": 40968,
    "location": "File > Export",
    "warning": "",
    # the following lines will be filled later on
    #"wiki_url": "",
    #"tracker_url": "https://sourceforge.net/projects/blenderextrainz/",
    "category": "Import-Export"}

import os
import sys
import time
import math
import datetime
import mathutils
import subprocess
import configparser

import bpy
import bpy.props
import bpy.utils

#### make a version number to do some unfortunatly necessary version switches
blender_version = (bpy.app.version[0] * pow(10, 6) +
                   bpy.app.version[1] * pow(10, 3) +
                   bpy.app.version[2])


#### global constants #########################################################

# state values
class STATUS:
    OK, WARNING, ERROR = range(3)


# log severities
class LOG:
    INFO, WARNING, ERROR, ADDINFO = range(4)


# attachement- und bone recognition for trainz
class PREFIX:
    ATTACHMENT = "a."
    BONE = "b.r."


# names for predefined vertex groups
class VG_NAME:
    ERROR_NO_MATERIAL_ASSIGNED = "ERROR_no_Material_assigned"
    ERROR_FACELESS_FACES = "ERROR_surfaceless_Polygons"
    WARNING_TO_MUCH_INFLUENCES = "WARNING_to_much_Influences"


# event-text definitions
class EVENT:
    TEXTBLOCK = 'events'
    TYPES = ['sound', 'soundsync', 'sync', '3dsys', 'generic',
             'effect', 'attach', 'detach', 'transfer', 'animcomplete',
             'loopcomplete', 'destroyed']


# event dictionary keys
class EVT:
    FRAME = 'f'
    TYPE = 'y'
    TRIGGER = 'r'


# recommended chars
class RECOMMENDED:
    CHARS = ['a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i', 'j', 'k', 'l', 'm',
             'n', 'o', 'p', 'q', 'r', 's', 't', 'u', 'v', 'w', 'x', 'y', 'z',
             '0', '1', '2', '3', '4', '5', '6', '7', '8', '9', '-', '_', 'A',
             'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N',
             'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', ':',
             '\\', '.', '/']


# desired selection method
class SELECTIONMETHOD:
    VISIBLE, SELECTED = ('visible', 'selected')


# desired error handling
class ERRORHANDLING:
    NONE, COLLECT, CORRECT = ('none', 'collect', 'correct')


# save dictionary keys
class SAFE:
    CURRENTFRAME = 'c'
    MODEACTIVEOBJECT = 'm'
    ACTIVEOBJECT = 'a'


# configuration
class CONFIG:
    selection_method = SELECTIONMETHOD.VISIBLE
    export_mesh = True
    export_animation = False
    export_scaled = False
    export_diffuse_as_ambient = True
    export_mirror_as_emit = False
    scaling_factor = 1.0
    write_log = True
    only_xml = False
    error_correction = ERRORHANDLING.COLLECT
    FILENAME = "export_trainz.cfg"
    LOGFILE_EXT = ".log"
    TMI_LOGFILE_EXT = "_TMI.log"
    XMLFILE_EXT = ".xml"


# option names
class OPTION:
    EXPORT_MESH = 'ExportMesh'
    EXPORT_ANIM = 'ExportAnimation'
    EXPORT_SCALED = 'ExportScaled'
    EXPORT_DIFFUSE_AS_AMBIENT = 'ExportDiffuseAsAmbient'
    EXPORT_MIRROR_AS_EMIT = 'ExportMirrorAsEmit'
    SCALING_FACTOR = 'ScalingFactor'
    WRITE_LOG = 'WriteLog'
    ONLY_XML = 'OnlyXML'
    ERROR_CORRECTION = 'ErrorCorrection'
    SELECTION_METHOD = 'SelectionMethod'


# config file
class CONFIGFILE:
    SECTION = 'DEFAULT'
    Parser = configparser.SafeConfigParser(
        {OPTION.SELECTION_METHOD: CONFIG.selection_method,
         OPTION.ERROR_CORRECTION: CONFIG.error_correction,
         OPTION.ONLY_XML: CONFIG.only_xml,
         OPTION.WRITE_LOG: CONFIG.write_log,
         OPTION.SCALING_FACTOR: CONFIG.scaling_factor,
         OPTION.EXPORT_MIRROR_AS_EMIT: CONFIG.export_mirror_as_emit,
         OPTION.EXPORT_DIFFUSE_AS_AMBIENT: CONFIG.export_diffuse_as_ambient,
         OPTION.EXPORT_SCALED: CONFIG.export_scaled,
         OPTION.EXPORT_ANIM: CONFIG.export_animation,
         OPTION.EXPORT_MESH: CONFIG.export_mesh})


# constants for floating point math
class FPM:
#    NDIGITS = 6  # digits after dot: 6 -> smallest possible value 1 Micrometer
    NDIGITS = 4  # digits after dot: 4 -> smallest possible value 1 Millimeter
    EPSILON = 0.5 * 10 ** - NDIGITS  # precision adjusted to decimal places
    LIMIT = 10 ** - NDIGITS  # distance within "double" vertices will welded


# material dictionary keys
class MAT:
    MATERIAL = 'm'
    DOUBLESIDED = 'd'
    #TRAINZPARENTS = 't'  # only needed for SKEL-based animation


# trainz_bone dictionary keys
class TB:
    CONTAINER = 'c'
    BONE = 'b'


# animation_basics dictionary keys
class AB:
    STARTFRAME = 's'
    ENDFRAME = 'e'
    FPS = 'f'


# indentations
IND1 = '  '
IND2 = IND1 + IND1
IND3 = IND1 + IND2
IND4 = IND1 + IND3
IND5 = IND1 + IND4
IND6 = IND1 + IND5


# object data dictionary keys
class OBJ:
    MAT = 'm'  # possibly scaled object matrix
    LOC = 'l'  # object location
    ROT = 'r'  # object rotation
    SCA = 's'  # object scale
    UVL = 'u'  # active uv-layer


# format strings
class STRINGF:
    VERTEX_PNT = ("<position>{co}</position>"
                  "<normal>{no}</normal>"
                  "<texcoord>{uv}</texcoord>")
    VERTEX_BB = ("<boneId stream=\"{s}\">{b}</boneId>"
                 "<blend stream=\"{s}\">"
                 "{w:." + str(FPM.NDIGITS) + "f}"
                 "</blend>")
    TRI_START = IND3 + "<triangle><materialId>%i</materialId>"
    VERTEX = "<vertex>%(p)s%(b)s</vertex>"
    TRI_END = "%s</triangle>\n"
    ATTACHMENT = IND3 + ("<attachment>"
                         "<name>{n}</name>"
                         "<position>{p}</position>"
                         "<orientation>{o}</orientation>"
                         "</attachment>\n")

    BONE = IND3 + ("<bone>"
                   "<id>%(i)i</id>"
                   "<name>%(n)s</name>"
                   "<parentName>%(tp)s</parentName>"
                   "<position>%(p)s</position>"
                   "<orientation>%(o)s</orientation>"
                   "</bone>\n")
    TEXTURE = IND4 + ("<texture>"
                      "<textureName>%(n)s</textureName>"
                      "<alphaSource>%(a)s</alphaSource>"
                      "<typeName>%(t)s</typeName>"
                      "<tileU>%(u)s</tileU>"
                      "<tileV>%(v)s</tileV>"
                      "<layerOpacity>%(o)f</layerOpacity>"
                      "%(f)s"
                      "</texture>\n")
    MATERIAL_PROPS = IND3 + ("<name>%(n)s</name>"
                             "<id>%(i)i</id>"
                             "<ambient>%(a)s</ambient>"
                             "<diffuse>%(d)s</diffuse>"
                             "<specular>%(s)s</specular>"
                             "<emissive>%(e)s</emissive>"
                             "<shine>%(h)f</shine>"
                             "<opacity>%(o)f</opacity>"
                             "<twoSided>%(t)s</twoSided>\n")
    ANIM_AND_TRACKS_OPENER = ("%(1)s<animations>\n"
                              "%(2)s<animation>\n"
                              "%(3)s<name></name>\n"
                              "%(3)s<frameRate>%(f)i</frameRate>\n"
                              "%(3)s<useLocalSpace>false</useLocalSpace>\n"
                              "%(3)s<animationTracks>\n")
    ANIMATIONTRACK_OPENER = ("%(4)s<animationTrack>\n"
                             "%(5)s<targetName>%(n)s</targetName>\n"
                             "%(5)s<keyFrames>\n")
    KEYFRAME = IND6 + ("<keyFrame>"
                       "<position>%(p)s</position>"
                       "<rotation>%(r)s</rotation>"
                       "</keyFrame>\n")
    EVENT = ("%(4)s<event>"
             "<frameIndex>%(f)i</frameIndex>"
             "<typeName>%(t)s</typeName>"
             "<eventName>%(e)s</eventName>"
             "</event>\n")
    F_TO_S = "{{:.{p}f}}".format(p=FPM.NDIGITS)  # to set a certain precision
    Q_TO_JQ = "{0}, {0}, {0}, {0}".format(F_TO_S)




# trainz texture type
class TTT:
    AMBIENT = "ambient"
    DIFFUSE = "diffuse"
    SPECULAR = "specular"
    SHINE = "shine"
    SHINESTRENGTH = "shinestrength"
    SELFILLUM = "selfillum"
    OPACITY = "opacity"
    FILTERCOLOR = "filtercolor"
    BUMP = "bump"
    REFLECT = "reflect"
    REFRACT = "refract"
    DISPLACEMENT = "displacement"


# material decorations
class DECO:
    DOT = "."
    MARKER = "m."
    NOTEX = "notex"
    ONETEX = "onetex"
    REFLECT = "reflect"
    GLOSS = "gloss"
    TBUMPTEX = "tbumptex"
    TBUMPGLOSS = "tbumpgloss"
    TBUMPENV = "tbumpenv"


# official material decorations
OFFICIAL_DECORATIONS = (DECO.NOTEX,
                        DECO.ONETEX,
                        DECO.REFLECT,
                        DECO.GLOSS,
                        DECO.TBUMPTEX,
                        DECO.TBUMPGLOSS,
                        DECO.TBUMPENV)


# examine script path
path = ''
for p in sys.path:
    if os.path.exists(p + "\\export_trainz.py"):
        path = p + "\\"


# store script path
class SCRIPT:
    PATH = path


# TMI stuff
class TMI:
    FILENAME = "TrainzMeshImporter.exe"


#### global functions #########################################################

def convert_forbidden_chars(s):
    ''' function to convert &<> into character entity references'''
    result = s.replace('&', '&amp;')
    result = result.replace('<', '&lt;')
    result = result.replace('>', '&gt;')
    return result


def compare_floats(a, b):
    '''function to compare floats against a predefined absolute epsilon'''
    if abs(a - b) < FPM.EPSILON:  # if difference < epsilon we assume equality
        result = 0
    else:  # a decision must be made which is the bigger one
        if a > b:
            result = 1
        else:
            result = -1
    return result


def compare_vert_locs(v1, v2):
    '''returns True if both Verts share the same position'''
    result = True
    for i in range(3):
        if compare_floats(v1.co[i], v2.co[i]) != 0:
            result = False
            break
    return result


def get_autosmooth_normal(mesh, fop, mesh_vi):
    '''check if smoothing has to be applied and return the depending normal'''
    ##  is at least one neighbourface smooth, the we return the vertex
    ##  normal(smooth applied), else the face normal will be returned
    result = fop.normal  # init with the normal for the un-smooth case

    ##  faulty (none-planar) faces may have a zero-length normal, and without
    ##  direction you can't calculate direction difference
    if fop.normal.length > 0.0:
        if blender_version < 2063000:
            for f in mesh.faces:
                if ((f != fop) and
                        (mesh_vi in f.vertices) and
                        (f.normal.length > 0.0)):
                    angle = int(round(
                        math.degrees(fop.normal.angle(f.normal))))
                    if angle <= mesh.auto_smooth_angle:
                        result = mesh.vertices[mesh_vi].normal

        else:
            for p in mesh.polygons:
                if ((p != fop) and
                        (mesh_vi in p.vertices) and
                        (p.normal.length > 0.0)):
                    angle = int(round(
                        math.degrees(fop.normal.angle(p.normal))))
                    if angle <= mesh.auto_smooth_angle:
                        result = mesh.vertices[mesh_vi].normal
    return result


def tupel_to_float_str(t):
    '''return the items of t as string of rounded values'''
    result = []
    for i in range(len(t)):
        result.append(STRINGF.F_TO_S.format(t[i]))
    return ', '.join(result)


def quat_to_jet_quat_str(q):
    '''return a Blender quaternion(w,x,y,z) rounded as
    quaternion string in Jet order(x,y,z,w)'''
    return STRINGF.Q_TO_JQ.format(q.x, q.y, q.z, q.w)


def add_fops_to_vertexgroup(group_name, objct, foplist):
    '''add vertices of FaceOrPolylist to vertexgroup group_name'''
    #print("VG_NAME.ERROR_NO_MATERIAL_ASSIGNED:\t",
    #      VG_NAME.ERROR_NO_MATERIAL_ASSIGNED)
    #print("group_name:\t", group_name)
    #print("objct:\t", objct)
    #print("foplist:\t", foplist)

    ## create or return the vertexgroup group_name
    try:
        vertex_group = objct.vertex_groups[group_name]
        #print("vertex group gab's schon")
    except:
        vertex_group = objct.vertex_groups.new(name=group_name)
        #print("vertex group neu angelegt")
    ## append vertices of polylist to found/created vertexgroup
    for fop in foplist:
        vertex_group.add(tuple(fop.vertices), 0.0, 'ADD')


def add_vertex_to_vertexgroup(group_name, objct, vertex_id):
    '''add vertices of facelist to vertexgroup group_name'''
    ## create or return the vertexgroup group_name
    try:
        vertex_group = objct.vertex_groups[group_name]
    except:
        vertex_group = objct.vertex_groups.new(name=group_name)
    ## append vertex_id to found/created vertexgroup
    vertex_group.add((vertex_id, ), 0.0, 'ADD')


#### class definitions ########################################################

class Error(Exception):
    '''Base class for exceptions in this module.'''

    def __init__(self, message):
        self.message = message

    def __str__(self):
        return self.message


class TrainzBoneItem(dict):
    '''class to distinguish trainz bones from other dicts (via type)'''
    pass


class TrainzExport:
    '''data and methods needed for export'''

    def __init__(self, filename, context):
        self.start_time = time.time()
        self.status = STATUS.OK
        self.export_filename = filename
        #print('self.export_filename: ', self.export_filename)
        self.log_filename = self.export_filename.replace(
            CONFIG.XMLFILE_EXT,
            CONFIG.LOGFILE_EXT)
        self.TMI_log_filename = self.export_filename.replace(
            CONFIG.XMLFILE_EXT,
            CONFIG.TMI_LOGFILE_EXT)
        #print('self.log_filename: ', self.log_filename)
        if os.path.exists(self.log_filename):  # delete previous log if found
            os.remove(self.log_filename)
        self.context = context
        self.meshes = list()  # list for meshes to export
        self.attachment_points = list()  # list for attachement points
        self.materials = list()  # list to store all materials used by meshes
        self.safe = dict()  # storage for values which may change during export
        self.trainz_bones = list()  # hold all bones
        self.root_bone = None  # store the root bone
        self.animation_basics = dict()
        self.events = list()  # list to store animation events

    def log(self, message, severity):
        '''print and/or log a message'''
        now = datetime.datetime.now()
        time_stamp = now.strftime('%H:%M:%S ')
        if severity == LOG.INFO:
            severity_text = 'INFO:    '
        elif severity == LOG.WARNING:
            if self.status != STATUS.ERROR:
                self.status = STATUS.WARNING
            severity_text = 'WARNING: '
        elif severity == LOG.ERROR:
            self.status = STATUS.ERROR
            severity_text = 'ERROR:   '
        elif severity == LOG.ADDINFO:
            severity_text = '   '
        print(time_stamp + severity_text + message)
        if CONFIG.write_log:
            with open(self.log_filename, mode='a', encoding='utf-8') as f:
                f.write(time_stamp + severity_text + message + "\n")
            f.close()

    def console_message(self, message, spaces=4):
        '''a centralized print function to print info messages'''
        now = datetime.datetime.now()
        time_stamp = now.strftime('%H:%M:%S')
        space = ' ' * spaces
        print(time_stamp + space + message)

    def save_state(self):
        '''saving values which may change during export'''
        # will be manipulated during anim export
        self.safe[SAFE.CURRENTFRAME] = self.context.scene.frame_current
        # store the current active object
        self.safe[SAFE.ACTIVEOBJECT] = self.context.scene.objects.active
        # objects in edit state return the properties before edit mode was
        # entered; so we leave edit mode and enter later again
        if ((self.context.scene.objects.active is not None) and
                self.context.scene.objects.active.is_visible(
                    self.context.scene)):
            self.safe[
                SAFE.MODEACTIVEOBJECT] = self.context.scene.objects.active.mode
            bpy.ops.object.mode_set(mode='OBJECT',
                                    toggle=False)

    def restore_state(self):
        '''restore previously saved values'''
        # restore active object
        self.context.scene.objects.active = self.safe[SAFE.ACTIVEOBJECT]
        # restore state of active object
        if ((self.context.scene.objects.active is not None) and
                self.context.scene.objects.active.is_visible(
                    self.context.scene)):
            bpy.ops.object.mode_set(mode=self.safe[SAFE.MODEACTIVEOBJECT],
                                    toggle=False)
        # restore currentFrame
        self.context.scene.frame_set(self.safe[SAFE.CURRENTFRAME])

    def get_mesh_name(self):
        '''extracts the mesh name from the blender filename'''
        s = os.path.basename(bpy.data.filepath).rsplit('.')[0]
        # drop all unrecommended chars in filename
        n = ''
        for i in range(len(s)):
            if s[i] in RECOMMENDED.CHARS:
                n += s[i]
        return n

    def get_material_id(self, material, double_sided):
        '''returns the index of material in self.Materials'''
        index = -1
        for i in range(len(self.materials)):
            if ((material == self.materials[i][MAT.MATERIAL]) and
                    (double_sided == self.materials[i][MAT.DOUBLESIDED])):
                index = i
                break
        return index

    def is_trainz_bone(self, bone):
        '''returns True if the naming convention for trainz bones are met'''
        return bone.name.find(PREFIX.BONE, 0, len(PREFIX.BONE)) != -1

    def get_trainz_bone_parent(self, obj):
        '''iterate through all parents until we find a member
        of self.trainz_bones or no furter parent exists'''
        ## initialise
        parent_bone = None
        trainz_bone = None
        #print('\nobj:\t', obj)
        #print('type(obj):\t', type(obj))
        ## "convert" obj to a TrainzBoneItem
        if isinstance(obj, TrainzBoneItem):
            # obj is a TrainzBoneItem
            trainz_bone = obj
        else:
            # looking for obj to be bone of a TrainzBoneItem
            for i in range(len(self.trainz_bones)):
                if obj == self.trainz_bones[i][TB.BONE]:
                    trainz_bone = self.trainz_bones[i]
                    break
            # if one of the subsequent parents is a trainz bone
            # we dirctly return them
            obj_parent = obj.parent
            while (obj_parent is not None) and (parent_bone is None):
                for i in range(len(self.trainz_bones)):
                    if obj_parent == self.trainz_bones[i][TB.BONE]:
                        parent_bone = self.trainz_bones[i][TB.BONE]
                        break
                obj_parent = obj_parent.parent
        ## if object was a trainzbone we search for its parent
        if trainz_bone is not None:
            ## iterate through TrainzBones for "contained" bones
            p = trainz_bone[TB.BONE].parent
            if trainz_bone[TB.CONTAINER] is not None:
                ## iterate through all PoseBones
                while (p is not None) and (parent_bone is None):
                    for i in range(len(self.trainz_bones)):
                        if p == self.trainz_bones[i][TB.BONE]:
                            parent_bone = p
                            break
                    if parent_bone is None:
                        p = p.parent
                # within the container no further trainz bone, so we jump out
                if parent_bone is None:
                    p = trainz_bone[TB.CONTAINER]
            ## iterate through TrainzBones for "free" bones (none-Posebones)
            while (p is not None) and (parent_bone is None):
                for i in range(len(self.trainz_bones)):
                    if p == self.trainz_bones[i][TB.BONE]:
                        parent_bone = p
                        break
                if parent_bone is None:
                    p = p.parent
        ## return result
        return parent_bone

    def remove_vertex_group(self, objct, vertex_group_name):
        '''removes an possibly in object existing vertex group'''
        ## check if the mentioned vertex group exists
        #print("attempt to delete vertex group ", vertex_group_name,
        #      " from object ", objct)
        try:
            remove_vg = objct.vertex_groups[vertex_group_name]
            delete = True
        except:
            delete = False
        ## delete existing vg
        if delete:
            # activate object of interrest
            self.context.scene.objects.active = objct
            # select the proper vg if not already done
            if objct.vertex_groups.active == remove_vg:
                reselect_vg = None
            else:
                reselect_vg = objct.vertex_groups.active
                bpy.ops.object.vertex_group_set_active(group=remove_vg.name)
            # delete active vertex group
            bpy.ops.object.vertex_group_remove()
            ## restore former selection
            if reselect_vg is not None:
                bpy.ops.object.vertex_group_set_active(group=reselect_vg.name)

    def get_vertex_pnt(self, obj_prop, mesh, face, face_vi):
        '''return position, normal and texcoords in XML format'''
        #### position
        ## multiplication order changed with version 2.59.0
        ## old order:
        if blender_version < 2059000:
            co = (obj_prop[OBJ.LOC] +
                  (mathutils.Vector([
                      mesh.vertices[face.vertices[face_vi]].co[0] *
                      obj_prop[OBJ.SCA][0],
                      mesh.vertices[face.vertices[face_vi]].co[1] *
                      obj_prop[OBJ.SCA][1],
                      mesh.vertices[face.vertices[face_vi]].co[2] *
                      obj_prop[OBJ.SCA][2]]) *
                   obj_prop[OBJ.ROT]))
        ## new order:
        else:
            co = (obj_prop[OBJ.LOC] +
                  mathutils.Vector(
                      obj_prop[OBJ.ROT] *
                      mathutils.Vector([
                          mesh.vertices[face.vertices[face_vi]].co[0] *
                          obj_prop[OBJ.SCA][0],
                          mesh.vertices[face.vertices[face_vi]].co[1] *
                          obj_prop[OBJ.SCA][1],
                          mesh.vertices[face.vertices[face_vi]].co[2] *
                          obj_prop[OBJ.SCA][2]])))

        #### normal
        ## multiplication order changed with version 2.59.0
        ## old order:
        if blender_version < 2059000:
            if face.use_smooth:
                if mesh.use_auto_smooth:
                    no = (get_autosmooth_normal(mesh,
                                                face,
                                                face.vertices[face_vi]) *
                          obj_prop[OBJ.ROT])
                else:
                    no = (mesh.vertices[face.vertices[face_vi]].normal *
                          obj_prop[OBJ.ROT])
            else:
                no = face.normal * obj_prop[OBJ.ROT]
        ## new order:
        else:
            if face.use_smooth:
                if mesh.use_auto_smooth:
                    no = mathutils.Vector(obj_prop[OBJ.ROT] *
                                          get_autosmooth_normal(
                                              mesh,
                                              face,
                                              face.vertices[face_vi]))
                else:
                    no = mathutils.Vector(obj_prop[OBJ.ROT] *
                                          mesh.vertices[
                                              face.vertices[face_vi]].normal)
            else:
                no = mathutils.Vector(obj_prop[OBJ.ROT] * face.normal)

        ## texture coords
        if obj_prop[OBJ.UVL] is not None:
            uv = (obj_prop[OBJ.UVL][face.index].uv[face_vi][0],
                  # v must be inverted for trainz
                  1 - obj_prop[OBJ.UVL][face.index].uv[face_vi][1])
        else:
            uv = (0.0, 0.0)

        ## hand over
        return STRINGF.VERTEX_PNT.format(co=tupel_to_float_str(co),
                                         no=tupel_to_float_str(no),
                                         uv=tupel_to_float_str(uv))

    def get_vertex_bb(self, bone_list, obj, mesh, face, vi):
        '''return the influences of a vertex in case of an animated mesh'''
        vertex_bb = []
        current_stream = 0
        #### we're looking for vertex groups named like trainz bones
        #### to add vg influences
        ## collect all influences to calculate the normalize-factor
        weight_sum = 0.0
        for vg in mesh.vertices[face.vertices[vi]].groups:
            for i, b in enumerate(bone_list):
                if obj.vertex_groups[vg.group].name == b[TB.BONE].name:
                    weight_sum += vg.weight
                    break
        if weight_sum > 0.0:
            normalize_factor = 1.0 / weight_sum
        else:
            normalize_factor = 1.0
        ## write out the vertexgroup influences normalized
        for vg in mesh.vertices[face.vertices[vi]].groups:
            for i, b in enumerate(bone_list):
                if obj.vertex_groups[vg.group].name == b[TB.BONE].name:
                    vertex_bb.append(STRINGF.VERTEX_BB.format(
                        s=current_stream,
                        b=i,
                        w=vg.weight * normalize_factor))
                    current_stream += 1
                    break
        ## if we have no vgs we search for a parental influence
        if current_stream == 0:
            ## search for a parent listed in our trainz-bone-list
            trainz_bone = None
            parent = obj.parent
            while (parent is not None) and (trainz_bone is None):
                if parent is not None:
                    for b in bone_list:
                        if parent == b[TB.BONE]:
                            trainz_bone = parent
                            break
                    if trainz_bone is None:
                        parent = parent.parent
            ## add parent bone influence if we found one
            if trainz_bone is not None:
                # get bone id
                for i, b in enumerate(bone_list):
                    if b[TB.BONE] == trainz_bone:
                        trainz_bone_id = i
                        break
                # the parent-influence has always a weight of 1 (100%)
                vertex_bb.append(STRINGF.VERTEX_BB.format(s=current_stream,
                                                          b=trainz_bone_id,
                                                          w=1.0))
        ## hand out result as string
        return ''.join(vertex_bb)

    def build_texture_node(self, tex_slot, trainz_tex_type, amount, sl):
        '''build the xml texture node'''

        # change from 0.95 to 0.96
        #alpha_source = str(bool(trainz_tex_type == TTT.OPACITY)).lower()
        alpha_source = str(bool((trainz_tex_type == TTT.OPACITY) or
                                tex_slot.texture.image.use_alpha)).lower()

        # tile: no distinction between u or v in Blender
        tile = str(tex_slot.texture.extension == 'REPEAT').lower()
        flip_green = ''
        if (trainz_tex_type == TTT.BUMP or
                trainz_tex_type == TTT.DISPLACEMENT):
            flip_green = "<flipGreen>false</flipGreen>"
        sl.append(STRINGF.TEXTURE % {
            'n': convert_forbidden_chars(bpy.path.abspath(
                tex_slot.texture.image.filepath)),
            'a': alpha_source,
            't': trainz_tex_type,
            'u': tile,
            'v': tile,
            'o': round(amount, FPM.NDIGITS),
            'f': flip_green})

    ################################ create ###################################
    def write_triangles(self, file):
        '''convert Blender faces to Trainz triangles and write them to file'''
        obj = {}  # a dict to pass this data "by reference" to the subroutines
        string_triangle = []  # list to collect all strings
        for objct in self.meshes:
            self.console_message("write triangles for " + objct.name + "...")
            obj_triangle_count = 0  # counting exported triangles per object
            if blender_version < 2063000:
                pass
            else:
                ## with blender 2.63 polygons replace faces -> faces have to
                ## be generated for rendering
                objct.data.calc_tessface()
            ## get the object matrix, extract translation, rotation,
            obj[OBJ.MAT] = objct.matrix_world.copy()
            ## scale if needed
            if CONFIG.export_scaled:
                obj[OBJ.MAT] = (objct.matrix_world.copy() *
                                CONFIG.scaling_factor)
            ## assign other needed properties
            obj[OBJ.LOC] = obj[OBJ.MAT].to_translation()
            obj[OBJ.ROT] = obj[OBJ.MAT].to_quaternion()
            obj[OBJ.SCA] = obj[OBJ.MAT].to_scale()
            obj[OBJ.UVL] = None
            if blender_version < 2063000:
                for uvt in objct.data.uv_textures:
                    if uvt.active_render:
                        obj[OBJ.UVL] = uvt.data
            else:
                for uvt in objct.data.tessface_uv_textures:
                    if uvt.active_render:
                        obj[OBJ.UVL] = uvt.data
            if blender_version < 2063000:
                ## iterate through all object faces
                for face in objct.data.faces:
                    ## export the first triangle
                    del string_triangle[:]
                    material_id = self.get_material_id(
                        objct.material_slots[face.material_index].material,
                        objct.data.show_double_sided)
                    string_triangle.append(STRINGF.TRI_START % material_id)
                    for vertex_id in (0, 1, 2):
                        vertex_pnt = self.get_vertex_pnt(obj,
                                                         objct.data,
                                                         face,
                                                         vertex_id)
                        vertex_bb = self.get_vertex_bb(self.trainz_bones,
                                                       objct,
                                                       objct.data,
                                                       face,
                                                       vertex_id)
                        string_triangle.append(STRINGF.VERTEX %
                                               {'p': vertex_pnt,
                                                'b': vertex_bb})
                    file.write(STRINGF.TRI_END % ''.join(string_triangle))
                    obj_triangle_count += 1
                    ## check for and export a second triangle
                    if len(face.vertices) == 4:
                        del string_triangle[:]
                        material_id = self.get_material_id(
                            objct.material_slots[face.material_index].material,
                            objct.data.show_double_sided)
                        string_triangle.append(STRINGF.TRI_START % material_id)
                        for vertex_id in (0, 2, 3):
                            vertex_pnt = self.get_vertex_pnt(obj,
                                                             objct.data,
                                                             face,
                                                             vertex_id)
                            vertex_bb = self.get_vertex_bb(self.trainz_bones,
                                                           objct,
                                                           objct.data,
                                                           face,
                                                           vertex_id)
                            string_triangle.append(STRINGF.VERTEX %
                                                   {'p': vertex_pnt,
                                                    'b': vertex_bb})
                        file.write(STRINGF.TRI_END % ''.join(string_triangle))
                        obj_triangle_count += 1
            else:
                ## iterate through all object faces;
                ## tesselation generates only triangles
                for face in objct.data.tessfaces:
                    del string_triangle[:]
                    material_id = self.get_material_id(
                        objct.material_slots[face.material_index].material,
                        objct.data.show_double_sided)
                    string_triangle.append(STRINGF.TRI_START % material_id)
                    for vertex_id in (0, 1, 2):
                        vertex_pnt = self.get_vertex_pnt(obj,
                                                         objct.data,
                                                         face,
                                                         vertex_id)
                        vertex_bb = self.get_vertex_bb(self.trainz_bones,
                                                       objct,
                                                       objct.data,
                                                       face,
                                                       vertex_id)
                        string_triangle.append(STRINGF.VERTEX %
                                               {'p': vertex_pnt,
                                                'b': vertex_bb})
                    file.write(STRINGF.TRI_END % ''.join(string_triangle))
                    obj_triangle_count += 1
                    ## check for and export a second triangle;
                    ## indeed tesselation generates only triangles,
                    ## but only Ngons were tesselated (and quads not)
                    if len(face.vertices) == 4:
                        del string_triangle[:]
                        material_id = self.get_material_id(
                            objct.material_slots[face.material_index].material,
                            objct.data.show_double_sided)
                        string_triangle.append(STRINGF.TRI_START % material_id)
                        for vertex_id in (0, 2, 3):
                            vertex_pnt = self.get_vertex_pnt(obj,
                                                             objct.data,
                                                             face,
                                                             vertex_id)
                            vertex_bb = self.get_vertex_bb(self.trainz_bones,
                                                           objct,
                                                           objct.data,
                                                           face,
                                                           vertex_id)
                            string_triangle.append(STRINGF.VERTEX %
                                                   {'p': vertex_pnt,
                                                    'b': vertex_bb})
                        file.write(STRINGF.TRI_END % ''.join(string_triangle))
                        obj_triangle_count += 1
            ## print out triangles per object   obj_triangle_count         
            self.console_message("   ...{:d} triangles written".format
                                 (obj_triangle_count))

    def write_attachments(self, file):
        '''write attachment-empties into file'''
        for ap in self.attachment_points:
            ap_name = ap.name.lower()
            #ap_name = ap.name <- in case we need :Cull instead of :cull
            ### throw a warning if the attachment name is changing
            if ap_name != ap.name:
                self.log("Empty \"{o}\" exported as \"{l}\"".format
                         (o=ap.name,
                          l=ap_name),
                         LOG.WARNING)
            ## get attachment point matrix
            ap_matrix = ap.matrix_world.copy()
            ## apply scaling if needed
            if CONFIG.export_scaled:
                ap_matrix = ap_matrix.copy() * CONFIG.scaling_factor
            ## write the current attachment point props
            file.write(
                STRINGF.ATTACHMENT.format(
                    n=convert_forbidden_chars(ap_name),
                    p=tupel_to_float_str(ap_matrix.to_translation()),
                    o=quat_to_jet_quat_str(ap_matrix.to_quaternion())))

    def write_mesh_section(self, file):
        '''create mesh section strings and write them into file'''
        file.write(IND1 + "<mesh>\n")  # mesh section opener
        file.write(IND2 + "<name>" +
                   convert_forbidden_chars(self.get_mesh_name()) +
                   "</name>\n")
        file.write(IND2 + "<triangles>\n")  # start triangle section
        self.write_triangles(file)  # convert Blender faces to Trainz triangles
        file.write(IND2 + "</triangles>\n")  # end triangle section
        ## append an attachment section if necessary
        if len(self.attachment_points) > 0:
            file.write(IND2 + "<attachments>\n")
            self.write_attachments(file)
            file.write(IND2 + "</attachments>\n")
        file.write(IND1 + "</mesh>\n")  # mesh section closer

    def write_skeleton_section(self, file):
        '''create skel section strings and write them into file'''
        if len(self.trainz_bones) > 0:
            ## skeleton section opener
            file.write(IND1 + "<skeleton>\n" + IND2 + "<bones>\n")
            for b_id, b in enumerate(self.trainz_bones):
                ## get the name of trainz parent bone
                trainz_parent_name = ''
                trainz_parent = self.get_trainz_bone_parent(b)
                if trainz_parent is not None:
                    trainz_parent_name = trainz_parent.name
                ## for armature bones "matrix" is used, that means the location
                ## of a bone is similar to its "head" property and NOT
                ## to its "tail" property
                if b[TB.CONTAINER] is None:
                    matrix = b[TB.BONE].matrix_world.copy()
                else:
                    matrix = (b[TB.CONTAINER].matrix_world.copy() *
                              b[TB.BONE].bone.matrix_local)
                ## apply scaling if desired
                if CONFIG.export_scaled:
                    matrix = matrix.copy() * CONFIG.scaling_factor
                ## split matrix
                position = matrix.to_translation()
                rotation = matrix.to_quaternion()
                ## append boneprops to xml
                file.write(
                    STRINGF.BONE % {
                        'i': b_id,
                        'n': convert_forbidden_chars(b[TB.BONE].name),
                        'tp': convert_forbidden_chars(trainz_parent_name),
                        'p': tupel_to_float_str(position),
                        'o': quat_to_jet_quat_str(rotation)})
            ## skeleton section closer
            file.write(IND2 + "</bones>\n" + IND1 + "</skeleton>\n")

    def write_material_section(self, file):
        '''dump the material definitions into file'''
        ## material section opener
        file.write(IND1 + '<materials>\n')
        ## iterate through all materials
        texture_sl = []  # use a string list for texture processing
        ambient = []  # list to hold rgb tupel
        diffuse = []  # list to hold rgb tupel
        specular = []  # list to hold rgb tupel
        emissive = []  # list to hold rgb tupel
        for mat_id, mat in enumerate(self.materials):
            ## material opener
            file.write(IND2 + '<material>\n')
            #### to check the material names, we need the textures first
            del texture_sl[:]
            ## textures
            #
            #  Jet name                    meaning               Blender name
            #
            #jTEX_AMBIENT       =  0 # -> ambient color         <- Amb(value)
            #jTEX_DIFFUSE       =  1 # -> diffuse color         <- Col
            #jTEX_SPECULAR      =  2 # -> specular color        <- Csp
            #jTEX_SHINE         =  3 # -> specular level        <- Spec(value)
            #jTEX_SHINESTRENGTH =  4 # -> glossiness            <- Hard(value)
            #jTEX_SELFILLUM     =  5 # -> self-illumination     <- Emit(value)
            #jTEX_OPACITY       =  6 # -> opacity               <- Alpha(value)
            #jTEX_FILTERCOLOR   =  7 # -> filter color          <- TransLu???
            #jTEX_BUMP          =  8 # -> normal                <- Nor(normal)
            #jTEX_REFLECT       =  9 # -> reflection            <- Ref(value)
            #jTEX_REFRACT       = 10 # -> refraction            <- RayMir???
            #jTEX_DISPLACEMENT  = 11 # -> displacement          <- Disp
            #
            ## to recognise the used textures we use some booleans
            ambient_texture_used = False
            diffuse_texture_used = False
            spec_col_texture_used = False
            spec_val_texture_used = False
            glossiness_texture_used = False
            self_illum_texture_used = False
            opacity_texture_used = False
            filter_color_texture_used = False
            normal_texture_used = False
            reflection_texture_used = False
            refraction_texture_used = False
            displacement_texture_used = False
            ## iterate through all enabled textures
            for i, ts in enumerate(mat[MAT.MATERIAL].texture_slots):
                if (mat[MAT.MATERIAL].use_textures[i] and
                        ts is not None and
                        ts.texture.type == 'IMAGE'):
                    if ts.use_map_ambient:
                        ambient_texture_used = True
                        self.build_texture_node(ts,
                                                TTT.AMBIENT,
                                                abs(ts.ambient_factor),
                                                texture_sl)  # "abs()"ed
                    if ts.use_map_color_diffuse and ts.texture_coords == 'UV':
                        diffuse_texture_used = True
                        self.build_texture_node(ts,
                                                TTT.DIFFUSE,
                                                ts.diffuse_color_factor,
                                                texture_sl)
                    if ts.use_map_color_spec:
                        spec_col_texture_used = True
                        self.build_texture_node(ts,
                                                TTT.SPECULAR,
                                                ts.specular_color_factor,
                                                texture_sl)
                    if ts.use_map_specular:
                        spec_val_texture_used = True
                        self.build_texture_node(ts,
                                                TTT.SHINE,
                                                abs(ts.specular_factor),
                                                texture_sl)  # "abs()"ed
                    if ts.use_map_hardness:
                        glossiness_texture_used = True
                        self.build_texture_node(ts,
                                                TTT.SHINESTRENGTH,
                                                abs(ts.hardness_factor),
                                                texture_sl)  # "abs()"ed
                    if ts.use_map_emit:
                        self_illum_texture_used = True
                        self.build_texture_node(ts,
                                                TTT.SELFILLUM,
                                                abs(ts.emit_factor),
                                                texture_sl)  # "abs()"ed
                    if ts.use_map_alpha:
                        opacity_texture_used = True
                        self.build_texture_node(ts,
                                                TTT.OPACITY,
                                                abs(ts.alpha_factor),
                                                texture_sl)  # "abs()"ed
                    if ts.use_map_translucency:
                        filter_color_texture_used = True
                        self.build_texture_node(ts,
                                                TTT.FILTERCOLOR,
                                                abs(ts.translucency_factor),
                                                texture_sl)  # "abs()"ed
                    if ts.use_map_normal:
                        normal_texture_used = True
                        self.build_texture_node(ts,
                                                TTT.BUMP,    # "abs()"ed and
                                                abs(ts.normal_factor) / 5.0,
                                                texture_sl)  # scaled
                    # use tex_coordinates to check for a reflection texture
                    if (ts.use_map_color_diffuse and
                            ts.texture_coords == 'REFLECTION'):
                        reflection_texture_used = True
                        # the reflection map influences the diffuse color,
                        # so we use diffuse_color_factor
                        self.build_texture_node(ts,
                                                TTT.REFLECT,
                                                ts.diffuse_color_factor,
                                                texture_sl)
                    # vvv ???? same as max refraction ???? vvv
                    #if ts.use_map_colortransmission:
                    #    refraction_texture_used = True
                    #    self.build_texture_node(ts, TTT.REFRACT, \
                    #        ts.colortransmission_factor, texture_sl)
                    if ts.use_map_displacement:
                        displacement_texture_used = True
                        self.build_texture_node(ts,
                                                TTT.DISPLACEMENT,
                                                abs(ts.displacement_factor),
                                                texture_sl)  # "abs()"ed
            ## as we know now what textures are used, we can
            ## decorate undecorated matnames and check decorated ones
            # only lower case characters allowed in material names
            material_name = mat[MAT.MATERIAL].name.lower()
            if material_name.find(DECO.MARKER) != -1:
                undecorated_material = False
                mat_decoration = (DECO.MARKER +
                                  material_name.split(DECO.MARKER)[-1])
                if (material_name.split(DECO.MARKER)[-1] not in
                        OFFICIAL_DECORATIONS):  # check for known mat decos
                    self.log("Material \"" + mat[MAT.MATERIAL].name +
                             "\" uses an unknown or inofficial Material "
                             "Decoration",
                             LOG.WARNING)
            else:
                undecorated_material = True
            ## choose material decoration according used textures
            # initiate, if no texture is applied(as NOTEX does cause
            # render errors by some asset types, traincar for example)
            mat_deco_proposal = DECO.MARKER + DECO.ONETEX
            ## material with unsupported single texture maps; as NOTEX does
            ## cause render errors by some asset types, traincar for example,
            ## I choose ONETEX instead
            if (ambient_texture_used or
                    spec_col_texture_used or
                    spec_val_texture_used or
                    glossiness_texture_used or
                    self_illum_texture_used or
                    opacity_texture_used or
                    filter_color_texture_used or
                    normal_texture_used or
                    refraction_texture_used or
                    displacement_texture_used):
                mat_deco_proposal = DECO.MARKER + DECO.ONETEX
            ## materials with one texture map
            if reflection_texture_used:
                mat_deco_proposal = DECO.MARKER + DECO.REFLECT
            if diffuse_texture_used:
                mat_deco_proposal = DECO.MARKER + DECO.ONETEX
            ## materials with two texture maps
            if diffuse_texture_used and opacity_texture_used:
                mat_deco_proposal = DECO.MARKER + DECO.ONETEX
            if diffuse_texture_used and reflection_texture_used:
                mat_deco_proposal = DECO.MARKER + DECO.REFLECT
            ## material with three texture maps
            if (diffuse_texture_used and
                    opacity_texture_used and
                    reflection_texture_used):
                mat_deco_proposal = DECO.MARKER + DECO.GLOSS
            ## materials with normal textures
            if normal_texture_used and diffuse_texture_used:
                if (opacity_texture_used and
                        reflection_texture_used):        # not supported
                    mat_deco_proposal = DECO.MARKER + DECO.TBUMPTEX
                    self.log("Material \"" + mat[MAT.MATERIAL] +
                             "\" uses Diffuse, Normal, Alpha and Reflection "
                             "Map; so far no Material Type support this, fall"
                             " back to tbumptex",
                             LOG.WARNING)
                if (reflection_texture_used and
                        (not opacity_texture_used)):     # gloss
                    mat_deco_proposal = DECO.MARKER + DECO.TBUMPGLOSS
                if ((not reflection_texture_used) and
                        opacity_texture_used):           # tex
                    mat_deco_proposal = DECO.MARKER + DECO.TBUMPTEX
                if ((not opacity_texture_used) and
                        (not reflection_texture_used)):  # tex or env
                    mat_deco_proposal = DECO.MARKER + DECO.TBUMPENV
            ## now assign or check against the proposal
            if undecorated_material:
                material_name += DECO.DOT + mat_deco_proposal
                self.log("Material \"" + mat[MAT.MATERIAL].name +
                         "\" exported as \"" + material_name + "\"",
                         LOG.INFO)
            else:
                ## inform if proposal differs from used material decoration
                if mat_decoration != mat_deco_proposal:
                    self.log("decoration for Material \"" +
                             mat[MAT.MATERIAL].name +
                             "\" differs from the proposed decoration \"" +
                             mat_deco_proposal + "\"",
                             LOG.INFO)
            #### material properties
            # to made it easier to set up a blender scene it's possible to
            # ignore the ambient value and export instead the diffuse value,
            # if you really export Blenders ambient value the appearance in
            # Trainz is totally different because the shaders used in Blender
            # and Trainz differs
            del ambient[:]
            if CONFIG.export_diffuse_as_ambient:
                ambient.append(mat[MAT.MATERIAL].diffuse_color.r *
                               mat[MAT.MATERIAL].diffuse_intensity)
                ambient.append(mat[MAT.MATERIAL].diffuse_color.g *
                               mat[MAT.MATERIAL].diffuse_intensity)
                ambient.append(mat[MAT.MATERIAL].diffuse_color.b *
                               mat[MAT.MATERIAL].diffuse_intensity)
            else:
                ambient.append(self.context.scene.world.ambient_color.r *
                               mat[MAT.MATERIAL].ambient)
                ambient.append(self.context.scene.world.ambient_color.g *
                               mat[MAT.MATERIAL].ambient)
                ambient.append(self.context.scene.world.ambient_color.b *
                               mat[MAT.MATERIAL].ambient)
            # diffuse_color will be combined with diffuse_intensity
            # to get the JetDiffuseColor
            del diffuse[:]
            diffuse.append(mat[MAT.MATERIAL].diffuse_color.r *
                           mat[MAT.MATERIAL].diffuse_intensity)
            diffuse.append(mat[MAT.MATERIAL].diffuse_color.g *
                           mat[MAT.MATERIAL].diffuse_intensity)
            diffuse.append(mat[MAT.MATERIAL].diffuse_color.b *
                           mat[MAT.MATERIAL].diffuse_intensity)
            # specular_color will be combined with specular_intensity to
            # get the JetSpecularColor
            del specular[:]
            specular.append(mat[MAT.MATERIAL].specular_color.r *
                            mat[MAT.MATERIAL].specular_intensity)
            specular.append(mat[MAT.MATERIAL].specular_color.g *
                            mat[MAT.MATERIAL].specular_intensity)
            specular.append(mat[MAT.MATERIAL].specular_color.b *
                            mat[MAT.MATERIAL].specular_intensity)
            # in blender the diffuse color is used for emitting light; to use
            # a different color like in max/gmax the "unused" mirror color in
            # the raytrace reflection section can be used but
            # without any visual feedback in Blender
            del emissive[:]
            if CONFIG.export_mirror_as_emit:
                emissive.append(mat[MAT.MATERIAL].mirror_color.r *
                                mat[MAT.MATERIAL].emit)
                emissive.append(mat[MAT.MATERIAL].mirror_color.g *
                                mat[MAT.MATERIAL].emit)
                emissive.append(mat[MAT.MATERIAL].mirror_color.b *
                                mat[MAT.MATERIAL].emit)
            else:
                emissive.append(mat[MAT.MATERIAL].diffuse_color.r *
                                mat[MAT.MATERIAL].emit)
                emissive.append(mat[MAT.MATERIAL].diffuse_color.g *
                                mat[MAT.MATERIAL].emit)
                emissive.append(mat[MAT.MATERIAL].diffuse_color.b *
                                mat[MAT.MATERIAL].emit)
            hardness = round((mat[MAT.MATERIAL].specular_hardness - 1) / 510.0,
                             FPM.NDIGITS)
            if mat[MAT.MATERIAL].use_transparency:
                opacity = round(mat[MAT.MATERIAL].alpha, FPM.NDIGITS)
            else:
                opacity = 1.0
            ## create & write the material property string
            file.write(
                STRINGF.MATERIAL_PROPS % {
                    'n': convert_forbidden_chars(material_name),
                    'i': mat_id,
                    'a': tupel_to_float_str(ambient),
                    'd': tupel_to_float_str(diffuse),
                    's': tupel_to_float_str(specular),
                    'e': tupel_to_float_str(emissive),
                    'h': hardness,
                    'o': opacity,
                    't': str(bool(mat[MAT.DOUBLESIDED])).lower()})
            ## now its time to create a texture section if needed
            if len(texture_sl) > 0:
                file.write(IND3 + "<textures>\n" + "".join(texture_sl) +
                           IND3 + "</textures>\n")
            ## material closer
            file.write(IND2 + "</material>\n")
        ## material section closer
        file.write(IND1 + "</materials>\n")

    def write_animation_section(self, file):
        '''translate animdata into trainz xml
           definitions and write them to file'''
        ## play animation and get all bone positions
        frames = []
        self.context.scene.frame_set(self.animation_basics[AB.STARTFRAME])
        while (self.context.scene.frame_current <=
               self.animation_basics[AB.ENDFRAME]):
            ## collect bone positions and orientations for current frame
            bones = {}
            for b in self.trainz_bones:
                if b[TB.CONTAINER] is None:
                    bone_matrix = b[TB.BONE].matrix_world.copy()
                else:
                    # PoseBones need to be multiplied with the container
                    # matrix and the rest matrix to get global coordinates
                    pb_matrix = (b[TB.BONE].bone.matrix_local.copy() *
                                 b[TB.BONE].matrix_basis)
                    bone_matrix = (b[TB.CONTAINER].matrix_world.copy() *
                                   pb_matrix)
                ## get bones matrix
                bones[b[TB.BONE]] = bone_matrix.copy()
                ## apply scaling if needed
                if CONFIG.export_scaled:
                    bones[b[TB.BONE]] = (bone_matrix.copy() *
                                         CONFIG.scaling_factor)
            ## append bonedict to framearray
            frames.append(bones)
            ## next frame, please
            self.context.scene.frame_set(self.context.scene.frame_current + 1)
        ## open animation section
        file.write(
            STRINGF.ANIM_AND_TRACKS_OPENER % {'1': IND1,
                                              '2': IND2,
                                              '3': IND3,
                                              'f': self.animation_basics[
                                                  AB.FPS]})
        ## build frame sequence for each bone
        for b in self.trainz_bones:
            ## bonetrack opener
            file.write(
                STRINGF.ANIMATIONTRACK_OPENER % {'4': IND4,
                                                 '5': IND5,
                                                 'n': convert_forbidden_chars(
                                                     b[TB.BONE].name)})
            ## drop keyframes for current bone
            for bones in frames:
                file.write(
                    STRINGF.KEYFRAME % {
                        'p': tupel_to_float_str(
                            bones[b[TB.BONE]].to_translation()),
                        'r': quat_to_jet_quat_str(
                            bones[b[TB.BONE]].to_quaternion())})
            ## bonetrack closer
            file.write(IND5 + "</keyFrames>\n" + IND4 + "</animationTrack>\n")
        ## animtracks closer
        file.write(IND3 + "</animationTracks>\n")
        ## write event section
        if len(self.events) > 0:
            ## event section opener
            file.write(IND3 + "<events>\n")
            for event in self.events:
                file.write(
                    STRINGF.EVENT % {'4': IND4,
                                     'f': event[EVT.FRAME],
                                     't': event[EVT.TYPE],
                                     'e': convert_forbidden_chars(
                                         event[EVT.TRIGGER])})
            ## event section closer
            file.write(IND3 + "</events>\n")
        ## close animation section
        file.write(IND2 + "</animation>\n" + IND1 + "</animations>\n")

    ############################# data collect ################################

    def get_meshes(self):
        '''collect mesh objects in current scene'''
        ## collect visible|selected
        if CONFIG.selection_method == SELECTIONMETHOD.VISIBLE:
            self.meshes = [ob for ob in self.context.scene.objects
                           if (ob.type == 'MESH' and
                               ob.is_visible(self.context.scene))]
        else:
            self.meshes = [ob for ob in self.context.scene.objects
                           if (ob.type == 'MESH' and
                               ob.select)]
        ## no collection? -> nothing to export!
        if len(self.meshes) == 0:
            self.log("no Mesh " + CONFIG.selection_method +
                     " - nothing to export",
                     LOG.ERROR)

    def get_attachment_points(self):
        '''collect empties named like attachment points'''
        ## collect visible|selected
        if CONFIG.selection_method == SELECTIONMETHOD.VISIBLE:
            self.attachment_points = [ob for ob in self.context.scene.objects
                                      if (ob.type == 'EMPTY' and
                                          ob.is_visible(self.context.scene))]
        else:
            self.attachment_points = [ob for ob in self.context.scene.objects
                                      if (ob.type == 'EMPTY' and
                                          ob.select)]
        ## drop all non-trainz empties
        for a in self.attachment_points[:]:
            if a.name.find(PREFIX.ATTACHMENT, 0, len(PREFIX.ATTACHMENT)) == -1:
                self.attachment_points.remove(a)

    def get_bones(self):
        '''collect lattices and armatures containing bones
           named like trainz bones'''
        ## first collect visible|selected lattices
        if CONFIG.selection_method == SELECTIONMETHOD.VISIBLE:
            lattices = [o for o in bpy.context.scene.objects
                        if (o.type == 'LATTICE' and
                            o.is_visible(self.context.scene) and
                            self.is_trainz_bone(o))]
        else:
            lattices = [o for o in bpy.context.scene.objects
                        if (o.type == 'LATTICE' and
                            o.select and
                            self.is_trainz_bone(o))]
        # append lattices to bone list
        for l in lattices:
            self.trainz_bones.append(TrainzBoneItem({TB.CONTAINER: None,
                                                     TB.BONE: l}))
        ## than collect pose bones and armatures
        if CONFIG.selection_method == SELECTIONMETHOD.VISIBLE:
            armatures = [o for o in bpy.context.scene.objects
                         if (o.type == 'ARMATURE' and
                             o.is_visible(self.context.scene))]
        else:
            armatures = [o for o in bpy.context.scene.objects
                         if o.type == 'ARMATURE' and o.select]
        # append armatures/pose bones to bone list
        for a in armatures:
            if self.is_trainz_bone(a):  # armatures ARE Trainz Bones ...
                self.trainz_bones.append(TrainzBoneItem({TB.CONTAINER: None,
                                                         TB.BONE: a}))
            for b in a.pose.bones:  # ... and/or can CONTAIN Trainz Bones
                if self.is_trainz_bone(b):
                    self.trainz_bones.append(TrainzBoneItem({TB.CONTAINER: a,
                                                             TB.BONE: b}))
        #print('\nget_bones result:')
        #for tb in self.trainz_bones:
        #    print('\t', tb)

    def check_meshes(self):
        '''mesh checking'''
        ## check for and handle invisible polygons
        faceless_fops = []
        cur_sel_verts = []
        for o in self.meshes:
            ## delete possibly leftover VG from previous check
            self.remove_vertex_group(o, VG_NAME.ERROR_FACELESS_FACES)
            ## now start checking
            m = o.data
            del faceless_fops[:]
            if blender_version < 2063000:
                for f in m.faces:
                    if not (f.normal.length > 0.0):
                        faceless_fops.append(f)
            else:
                for p in m.polygons:
                    if not (p.normal.length > 0.0):
                        faceless_fops.append(p)
            if len(faceless_fops) > 0:
                self.log("surfaceless Polygon(s) in Object \"" + o.name +
                         "\", Mesh \"" + m.name + "\" detected",
                         LOG.WARNING)
                if CONFIG.error_correction == ERRORHANDLING.COLLECT:
                    add_fops_to_vertexgroup(VG_NAME.ERROR_FACELESS_FACES,
                                            o,
                                            faceless_fops)
                    self.log("Vertices of surfaceless Polygon(s) gathered in"
                             " Vertex Group \"" +
                             VG_NAME.ERROR_FACELESS_FACES + "\"",
                             LOG.INFO)
                elif CONFIG.error_correction == ERRORHANDLING.CORRECT:
                    ## memorize the current vertex selection and unselect all
                    del cur_sel_verts[:]
                    for i, v in enumerate(m.vertices):
                        if v.select:
                            cur_sel_verts.append(m.vertices[i])
                        m.vertices[i].select = False
                    try:
                        ## select all vertices related to facless faces and
                        ## call "remove doubles" to eliminate them;
                        ## works with faces and polygons
                        for fop in faceless_fops:
                            for i in fop.vertices:
                                m.vertices[i].select = True
                        ## operators only work with the active object,
                        ## so we must activate the desired one
                        self.context.scene.objects.active = o
                        bpy.ops.object.mode_set(mode='EDIT', toggle=False)
                        try:
                            if blender_version < 2063000:
                                bpy.ops.mesh.remove_doubles(
                                    limit=FPM.LIMIT)
                            else:
                                bpy.ops.mesh.remove_doubles(
                                    mergedist=FPM.LIMIT)
                            self.log('surfaceless Polygon(s) removed',
                                     LOG.INFO)
                        finally:
                            bpy.ops.mesh.select_all(action='DESELECT')
                            bpy.ops.object.mode_set(mode='OBJECT',
                                                    toggle=False)
                    finally:
                        ## restore the former selection
                        for i, v1 in enumerate(m.vertices):
                            for v2 in cur_sel_verts:
                                if ((not m.vertices[i].select) and
                                        compare_vert_locs(v1, v2)):
                                    m.vertices[i].select = True

        ## check if all objects using UV-mapping also have UV-coordinates
        error_messages = []
        for o in self.meshes:
            if o.data.uv_textures.active is None:
                for ms in o.material_slots:
                    for tm in self.materials:
                        # we check only if the material is getting exported
                        if tm[MAT.MATERIAL] == ms.material:
                            for i, ts in enumerate(ms.material.texture_slots):
                                if (ms.material.use_textures[i] and
                                        ts is not None and
                                        ts.texture_coords == 'UV' and
                                        ts.texture.type == 'IMAGE'):
                                    error_messages.append(
                                        "no UV-Layer for UV-Mapped Object \"" +
                                        o.name + "\", Mesh \"" + o.data.name +
                                        "\"")
        if len(error_messages) > 0:
            error_messages.sort()
            cur_mes = error_messages[0]
            i = 1
            while i < len(error_messages):
                if error_messages[i] == cur_mes:
                    error_messages.remove(cur_mes)
                else:
                    cur_mes = error_messages[i]
                    i += 1
            for em in error_messages:
                self.log(em, LOG.ERROR)
        ## inform about double sided objects
        for o in self.meshes:
            next_please = False
            if o.data.show_double_sided:
                for ms in o.material_slots:
                    for tm in self.materials:
                        # we complian only if the material is getting exported
                        if tm[MAT.MATERIAL] == ms.material:
                            self.log("Object \"" + o.name + "\" will be "
                                     "exported with Double Sided Faces",
                                     LOG.INFO)
                            next_please = True
                            break
                    if next_please:
                        break

    def check_hierarchy(self):
        '''if bones in use all meshes need to be part of the hierarchy'''
        ## if we have bones all meshes must be "connected" to the skeleton
        if self.root_bone is not None:
            for o in self.meshes:
                p_bone = None
                p = o.parent
                #iterate until parent bone or no (further) parent
                while (p is not None) and (p_bone is None):
                    if p is not None:
                        for tb in self.trainz_bones:
                            if p == tb[TB.BONE]:
                                p_bone = p
                                break
                        if p_bone is None:
                            p = p.parent
                if p is None:
                    self.log('if Trainz Bones shall be exported, Object "' +
                             o.name + '" must have a parent',
                             LOG.ERROR)
                elif p_bone is None:
                    self.log("the parent list for " + o.type.capitalize() +
                             " \"" + o.name + "\" must include at least one "
                             "Trainz Bone",
                             LOG.ERROR)

    def get_materials(self):
        '''collect all materials used in self.meshes'''
        # list to count how often a material index is used
        material_index_usage = []
        # list to collect all faces using a given material index
        faces_using_index = []
        # dict containing the material, if its a double sided mesh(is a
        # material information inside trainz) and the object parent(needed
        # info for SKEL-based animation)
        material_item = {}

        ## iterate through all meshes to export
        for o in self.meshes:
            ## delete possibly leftover VG from previous run
            self.remove_vertex_group(o, VG_NAME.ERROR_NO_MATERIAL_ASSIGNED)
            ## init used lists
            del material_index_usage[:]
            del faces_using_index[:]
            ## check if at least one material is used
            if len(o.material_slots) > 0:
                for i in range(len(o.material_slots)):  # init lists
                    material_index_usage.append(0)
                    faces_using_index.append(list())
                mesh = o.data
                if blender_version < 2063000:  # pre 2.63 code
                    for f in mesh.faces:
                        material_index_usage[f.material_index] += 1
                        faces_using_index[f.material_index].append(f)
                else:  # faces are replaced by polygons from 2.63 on
                    for p in mesh.polygons:
                        material_index_usage[p.material_index] += 1
                        faces_using_index[p.material_index].append(p)
            #print('o:', o)
            #print('material_index:', material_index)
            #print('faces_using_index:', faces_using_index)
            ## check if object has at least one material assigned
            if len(material_index_usage) == 0:
                self.log("Object \"" + o.name + "\" has no material assigned.",
                         LOG.ERROR)
            ## check if we have single faces without an assigned material
            else:
                for i in range(len(material_index_usage)):
                    if ((o.material_slots[i].material is None) and
                            (len(faces_using_index[i]) > 0)):
                        add_fops_to_vertexgroup(
                            VG_NAME.ERROR_NO_MATERIAL_ASSIGNED,
                            o,
                            faces_using_index[i])
                        self.log("Face(s) without Material in Object \"" +
                                 o.name + "\", Mesh \"" + mesh.name +
                                 "\" detected. Collect them in Vertex "
                                 "Group \"" +
                                 VG_NAME.ERROR_NO_MATERIAL_ASSIGNED + "\"",
                                 LOG.ERROR)
                    else:
                        if len(faces_using_index[i]) > 0:
                            ## create material_item
                            material_item[
                                MAT.MATERIAL] = o.material_slots[i].material
                            material_item[
                                MAT.DOUBLESIDED] = o.data.show_double_sided
                            ## add material_item to material list if necessary
                            append_material_item = True
                            for i in range(len(self.materials)):
                                if (self.materials[i][MAT.MATERIAL] ==
                                        material_item[MAT.MATERIAL] and
                                        self.materials[i][MAT.DOUBLESIDED] ==
                                        material_item[MAT.DOUBLESIDED]):
                                    append_material_item = False
                                    #self.materials[i][MAT.TRAINZPARENTS].add(
                                    #    self.get_trainz_bone_parent(o))
                                    break
                            if append_material_item:
                                #material_item[MAT.TRAINZPARENTS] = set(
                                #    [self.get_trainz_bone_parent(o)])
                                self.materials.append(material_item.copy())
        #print('\nget_materials result:')
        #for m in self.materials:
        #    print('\t', m)

    def log_message_list(self, messages, severity):
        '''remove duplicates and log messages'''
        if len(messages) > 0:
            messages.sort()
            cur_mes = messages[0]
            i = 1
            while i < len(messages):
                if messages[i] == cur_mes:
                    messages.remove(cur_mes)
                else:
                    cur_mes = messages[i]
                    i += 1
            for m in messages:
                self.log(m, severity)

    def check_materials(self):
        '''check if all self.Materials are Trainz compatible'''
        # material checks are realy difficult without materials
        if len(self.materials) > 0:
            ## check if material names only differ in lower/upper case
            ## (and ignore cases related to double/single sided faces)
            i = 0
            while i < len(self.materials):
                cur_mat = self.materials[i]
                for j in range(i + 1, len(self.materials)):
                    if (cur_mat[MAT.MATERIAL].name.lower() ==
                            self.materials[j][MAT.MATERIAL].name.lower() and
                            cur_mat[MAT.MATERIAL] !=
                            self.materials[j][MAT.MATERIAL]):
                        self.log("Material names \"%(m1)s\" and \"%(m2)s\" are"
                                 " identical if compared case-insensitive" % {
                                     'm1': cur_mat[MAT.MATERIAL].name,
                                     'm2': self.materials[j][
                                         MAT.MATERIAL].name},
                                 LOG.ERROR)
                i += 1
            ## check if a material is used for single AND double sided faces
            i = 0
            while i < len(self.materials):
                cur_mat = self.materials[i]
                outm = []  # list for Objects Using This Material
                for j in range(i + 1, len(self.materials)):
                    if (cur_mat[MAT.MATERIAL] == self.materials[j][
                            MAT.MATERIAL] and
                        cur_mat[MAT.DOUBLESIDED] != self.materials[j][
                            MAT.DOUBLESIDED]):
                        # collect involved objects
                        del outm[:]
                        for o in self.meshes:
                            for ms in o.material_slots:
                                if ms.material == cur_mat[MAT.MATERIAL]:
                                    outm.append(o)
                                    break
                        oss = [o for o in outm if not o.data.show_double_sided]
                        ods = [o for o in outm if o.data.show_double_sided]
                        ss_objects = ''
                        for o in oss:
                            ss_objects += '"' + o.name + '", '
                        ss_objects = ss_objects[:-2]
                        ds_objects = ''
                        for o in ods:
                            ds_objects += '"' + o.name + '", '
                        ds_objects = ds_objects[:-2]
                        self.log("Material \"%(m)s\" is used for Single AND "
                                 "Double Sided Objects; Single Sided: "
                                 "%(s)s, Double Sided: %(d)s" % {
                                     'm': cur_mat[MAT.MATERIAL].name,
                                     's': ss_objects,
                                     'd': ds_objects},
                                 LOG.ERROR)
                i += 1
            ## check if all materials use only textures of type image and
            ## if those image exists
            error_messages = []
            warning_messages = []
            for m in self.materials:
                tex_names_generic = []  # procedural textures used
                tex_names_missing = []  # assigned image not found
                # change from 0.95 to 0.96: warn about textures which
                # forced to use alpha but have no alpha plane
                tex_names_noplane = []  # no alpha plane but use_alpha checked

                for i, ts in enumerate(m[MAT.MATERIAL].texture_slots):
                    if m[MAT.MATERIAL].use_textures[i] and (ts is not None):
                        if (ts.texture.type != 'NONE' and
                                ts.texture.type != 'IMAGE'):
                            tex_names_generic.append(ts.texture.name)
                        elif ts.texture.type == 'IMAGE':
                            if ((ts.texture.image is None) or
                                not os.path.exists(bpy.path.abspath(
                                    ts.texture.image.filepath))):
                                tex_names_missing.append(ts.texture.name)
                            if ((ts.texture.image is not None) and
                                (ts.texture.image.use_alpha and
                                 (ts.texture.image.depth == 24))):
                                tex_names_noplane.append(ts.texture.name)

                if len(tex_names_generic) > 0:
                    error_messages.append(
                        "procedural texture(s) assigned to Material \"%(m)s\","
                        " Texture(s) \"%(t)s\"" % {
                            'm': m[MAT.MATERIAL].name,
                            't': ', '.join(tex_names_generic)})
                if len(tex_names_missing) > 0:
                    error_messages.append(
                        "image file(s) not found for Material \"%(m)s\", "
                        "Texture(s) \"%(t)s\"" % {
                            'm': m[MAT.MATERIAL].name,
                            't': ', '.join(tex_names_missing)})
                if len(tex_names_noplane) > 0:
                    warning_messages.append(
                        "no alpha plane in image(s) for Material \"%(m)s\", "
                        "Texture(s) \"%(t)s\"" % {
                            'm': m[MAT.MATERIAL].name,
                            't': ', '.join(tex_names_noplane)})
            self.log_message_list(error_messages, LOG.ERROR)
            self.log_message_list(warning_messages, LOG.WARNING)

            ## check if all textures mapped to alpha, specular and
            ## normal are uv-mapped
            error_messages = []
            for m in self.materials:
                tex_names = []
                for i, ts in enumerate(m[MAT.MATERIAL].texture_slots):
                    if m[MAT.MATERIAL].use_textures[i] and (ts is not None):
                        if (ts.texture.type == 'IMAGE' and
                                ts.use_map_alpha or
                                ts.use_map_specular or
                                ts.use_map_normal):
                            if ts.texture_coords != 'UV':
                                tex_names.append(ts.texture.name)
                if len(tex_names) > 0:
                    error_messages.append(
                        "UV mapping needed for Materials \"%(m)s\", "
                        "Textures(s) \"%(t)s\"" % {
                            'm': m[MAT.MATERIAL].name,
                            't': ', '.join(tex_names)})
            self.log_message_list(error_messages, LOG.ERROR)
            ## check if all textures maped to diffuse are uv- or
            ## reflection-mapped
            error_messages = []
            for m in self.materials:
                tex_names = []
                for i, ts in enumerate(m[MAT.MATERIAL].texture_slots):
                    if m[MAT.MATERIAL].use_textures[i] and (ts is not None):
                        if (ts.texture.type == 'IMAGE' and
                                ts.use_map_color_diffuse):
                            if (ts.texture_coords != 'UV' and
                                    ts.texture_coords != 'REFLECTION'):
                                tex_names.append(ts.texture.name)
                if len(tex_names) > 0:
                    error_messages.append(
                        "UV or Reflection Mapping needed for Material \"%(m)s"
                        "\", Textures(s) \"%(t)s\"" % {
                            'm': m[MAT.MATERIAL].name,
                            't': ', '.join(tex_names)})
            self.log_message_list(error_messages, LOG.ERROR)

    def get_and_check_root_bone(self):
        '''go for the root bone and drop a warning if we have a problem'''
        ## check bone hierarchie
        root_bones = []
        bone_names = []
        for tb in self.trainz_bones:
            p = self.get_trainz_bone_parent(tb)
            if p is None:
                root_bones.append(tb)
        #print('root_bones: ', root_bones)
        ## only one root bone is allowed
        if len(root_bones) == 1:
            self.root_bone = root_bones[0]
        elif len(root_bones) > 1:
            ## collect names of all root - bones
            del bone_names[:]
            for rb in root_bones:
                if rb[TB.CONTAINER] is None:
                    bone_names.append(rb[TB.BONE].type.capitalize() +
                                      " \"" + rb[TB.BONE].name + "\"")
                else:
                    bone_names.append(rb[TB.CONTAINER].type.capitalize() +
                                      " \"" + rb[TB.CONTAINER].name +
                                      "\" Bone \"" + rb[TB.BONE].name + "\"")
            ## propagate te problem
            self.log("found more than one Trainz Root Bone: " +
                     ", ".join(bone_names),
                     LOG.ERROR)
        ## the following checks are only necessary if we have the one root bone
        if self.root_bone is not None:
            ## check that the root bone is not rotated or
            ## moved(MAY lead to strange results in Trainz
            ## if BOTH happen, therefore only a warning)
            if self.root_bone[TB.CONTAINER] is not None:
                loc_x, loc_y, loc_z = (
                    self.root_bone[
                        TB.CONTAINER].matrix_world.to_translation() +
                    self.root_bone[TB.BONE].matrix.to_translation())
                rot_w, rot_x, rot_y, rot_z = (
                    self.root_bone[TB.BONE].matrix.to_quaternion().cross(
                        self.root_bone[
                            TB.CONTAINER].matrix_world.to_quaternion()))
            else:
                loc_x, loc_y, loc_z = (
                    self.root_bone[TB.BONE].matrix_world.to_translation())
                rot_w, rot_x, rot_y, rot_z = (
                    self.root_bone[TB.BONE].matrix_world.to_quaternion())
            rb_moved = (compare_floats(loc_x, 0.0) != 0 or
                        compare_floats(loc_y, 0.0) != 0 or
                        compare_floats(loc_z, 0.0) != 0)
            rb_rotated = (compare_floats(rot_w, 1.0) != 0 or
                          compare_floats(rot_x, 0.0) != 0 or
                          compare_floats(rot_y, 0.0) != 0 or
                          compare_floats(rot_z, 0.0) != 0)
            if rb_moved and rb_rotated:
                rb_name = self.root_bone[TB.BONE].name
                if self.root_bone[TB.CONTAINER] is not None:
                    rb_name = (
                        self.root_bone[TB.CONTAINER].name + '->' + rb_name)
                self.log("The Trainz Root Bone \"" + rb_name +
                         "\" is rotated and not located at the point of "
                         "origin. In Trainz your Object might not appear "
                         "where you expect them.",
                         LOG.WARNING)

    def check_influence(self):
        '''check if vertices are influenced by up to 4 vertex groups
           (trainz maximum is 4 streams) and are normalized'''
        ## collect all bone names
        bone_names = []
        influence_group_names = []
        influence_groups = []
        vertexgroup_created = False
        for bone in self.trainz_bones:
            bone_names.append(bone[TB.BONE].name)
        ## iterate trough all meshes and its vertices and
        ## count vertexgroups with bone names for every object
        for objct in self.meshes:
            ## delete possibly leftover VG from previous check
            self.remove_vertex_group(objct, VG_NAME.WARNING_TO_MUCH_INFLUENCES)
            ## start counting
            mesh = objct.data
            for vertex in mesh.vertices:
                ## extract all influence groups and check for
                del influence_group_names[:]
                del influence_groups[:]
                for group in vertex.groups:
                    if objct.vertex_groups[group.group].name in bone_names:
                        influence_group_names.append(
                            objct.vertex_groups[group.group].name)
                        influence_groups.append(
                            objct.vertex_groups[group.group])
                ## a) not more than 4 trainz bone groups groups are allowed
                if len(influence_group_names) > 4:
                    # show message
                    text = ("Vertex %(v)i in Mesh \"%(m)s\" is influenced by "
                            "more than 4 vertex groups:\n\t\t\tObject "
                            "\"%(o)s\", Vertex Group \"" % {
                                'v': vertex.index,
                                'm': mesh.name,
                                'o': objct.name} +
                            ("\"\n\t\t\tObject \"%s\",  Vertex "
                             "Group \"" % objct.name).join(
                                 influence_group_names))
                    self.log(text, LOG.WARNING)
                    # add vertex to VG "overinfluenced"
                    add_vertex_to_vertexgroup(
                        VG_NAME.WARNING_TO_MUCH_INFLUENCES,
                        objct,
                        vertex.index)
                    vertexgroup_created = True
            if vertexgroup_created:
                self.log("Vertices with to many influences gathered in "
                         "Vertex Group \"" +
                         VG_NAME.WARNING_TO_MUCH_INFLUENCES +
                         "\" of Object \"" + objct.name + "\", Mesh \"" +
                         mesh.name + "\"",
                         LOG.INFO)
                vertexgroup_created = False

    def get_animation_basics(self):
        '''collect basic informations needed to export animations'''
        ## check preconditions; we do this only
        ## if animation export is ordered
        if CONFIG.export_animation:
            # check if bones exist
            if len(self.trainz_bones) == 0:
                CONFIG.export_animation = False
                self.log("no Trainz Bones shall be exported, "
                         "animation export skipped",
                         LOG.WARNING)
            else:

                ## check if exported trainz bones are animated
                #CONFIG.export_animation = False
                #for b in self.trainz_bones:
                #    if b[TB.CONTAINER] is None:
                #        CONFIG.export_animation = (
                #            b[TB.BONE].animation_data is not None)
                #    else:
                #        CONFIG.export_animation = (
                #            b[TB.CONTAINER].animation_data is not None)
                #    if CONFIG.export_animation:
                #        break

                # check if we have at least one keyframe_point or a
                # driver in one of this scenes objects; the movement of
                # trainz bones can be inherited or "constrained" by them
                CONFIG.export_animation = False
                for o in self.context.scene.objects:
                    if o.animation_data is not None:
                        if o.animation_data.action is not None:
                            for f in o.animation_data.action.fcurves:
                                if len(f.keyframe_points) > 0:
                                    CONFIG.export_animation = True
                        if len(o.animation_data.drivers) > 0:
                            CONFIG.export_animation = True
                # warn if we have reset an requested animation export
                if not CONFIG.export_animation:
                    self.log("no animation data found, "
                             "animation export skipped",
                             LOG.WARNING)
        # if all preconditions checked we may collect animdatas ;)
        if CONFIG.export_animation:
            self.animation_basics[AB.STARTFRAME] = (
                self.context.scene.frame_start)
            self.animation_basics[AB.ENDFRAME] = (
                self.context.scene.frame_end)
            self.animation_basics[AB.FPS] = (
                self.context.scene.render.fps)

    def get_animation_events(self):
        '''collect events associated with specific frames'''
        # search for an event text block
        event_text = None
        event = {}
        for t in bpy.data.texts:
            if t.name.lower() == EVENT.TEXTBLOCK:
                event_text = t
                break
        if event_text is None:
            if CONFIG.export_animation:
                self.log("no event text found, no events generated", LOG.INFO)
        else:
            for i, line in enumerate(t.lines):
                if len(line.body) > 0:
                    values = line.body.split(None, 2)
                    if len(values) < 3:
                        self.log("ignore event \"" + line.body +
                                 "\" at line " + str(i + 1) + ": wrong format",
                                 LOG.WARNING)
                    else:
                        event.clear()
                        event[EVT.FRAME] = int(values[0])
                        # to stay CCG conform we allow
                        # usage of strings like 'Sound_Event':
                        event[EVT.TYPE] = values[1].lower().split('_')[0]
                        event[EVT.TRIGGER] = values[2]
                        highest_frame = (
                            self.animation_basics[AB.ENDFRAME] -
                            self.animation_basics[AB.STARTFRAME])
                        if event[EVT.FRAME] < 0:
                            self.log("ignore event \"" + line.body +
                                     "\" at line " + str(i + 1) +
                                     ": framenumber lower than first frame(0)",
                                     LOG.WARNING)
                        elif event[EVT.FRAME] > highest_frame:
                            self.log("ignore event \"" + line.body +
                                     "\" at line " + str(i + 1) +
                                     ": framenumber higher than last frame(" +
                                     str(highest_frame) + ")",
                                     LOG.WARNING)
                        else:
                            if event[EVT.TYPE] not in EVENT.TYPES:
                                self.log("ignore event \"" + line.body +
                                         "\" at line " + str(i + 1) +
                                         ": unknown event type",
                                         LOG.WARNING)
                            else:
                                self.events.append(event.copy())

    def check_for_unrecommended_characters(self, s):
        '''return a list of all unrecommended characters found in s'''
        complained_chars = []
        for c in s:
            if c not in RECOMMENDED.CHARS:
                complained_chars.append(c)
        return ''.join(complained_chars)

    def check_names(self):
        '''check all exported names for unrecommended chars'''
        ## attachment names
        for a in self.attachment_points:
            complained_chars = self.check_for_unrecommended_characters(a.name)
            if len(complained_chars) > 0:
                self.log("Empty \"" + a.name +
                         "\" contains unrecommended characters: " +
                         complained_chars,
                         LOG.INFO)
        ## material names
        for m in self.materials:
            complained_chars = (
                self.check_for_unrecommended_characters(m[MAT.MATERIAL].name))
            if len(complained_chars) > 0:
                self.log("Material \"" + m[MAT.MATERIAL].name +
                         "\" contains unrecommended characters: " +
                         complained_chars,
                         LOG.INFO)
        ## texture paths
        for m in self.materials:
            for i, ts in enumerate(m[MAT.MATERIAL].texture_slots):
                if (m[MAT.MATERIAL].use_textures[i] and
                        ts is not None and
                        ts.texture.type == 'IMAGE' and
                        ts.texture.image is not None):
                    complained_chars = (
                        self.check_for_unrecommended_characters(
                            ts.texture.image.filepath))
                    if len(complained_chars) > 0:
                        self.log("texture path \"" +
                                 ts.texture.image.filepath +
                                 "\" contains unrecommended characters: " +
                                 complained_chars,
                                 LOG.INFO)
        ## bone names
        for b in self.trainz_bones:
            complained_chars = (
                self.check_for_unrecommended_characters(b[TB.BONE].name))
            if len(complained_chars) > 0:
                if b[TB.CONTAINER] is None:
                    text = (b[TB.BONE].type.capitalize() + " \"" +
                            b[TB.BONE].name +
                            "\" contains unrecommended characters: " +
                            complained_chars)
                else:
                    text = ("Bone \"" + b[TB.BONE].name + "\" in Armature \"" +
                            b[TB.CONTAINER].name +
                            "\" contains unrecommended characters: " +
                            complained_chars)
                self.log(text, LOG.INFO)
        ## event names
        for e in self.events:
            complained_chars = (
                self.check_for_unrecommended_characters(e[EVT.TRIGGER]))
            if len(complained_chars) > 0:
                self.log("Event trigger \"" + e[EVT.TRIGGER] +
                         "\" contains unrecommended characters: " +
                         complained_chars,
                         LOG.INFO)

    def collect_data(self):
        '''collect and evaluate all data needed to
            export the visible object(s)'''
        ## print exporter settings
        self.log("", LOG.ADDINFO)
        self.log(OPTION.SELECTION_METHOD + ":\t\t" +
                 str(CONFIG.selection_method).capitalize(),
                 LOG.ADDINFO)
        self.log(OPTION.EXPORT_MESH + ":\t\t\t" + str(CONFIG.export_mesh),
                 LOG.ADDINFO)
        self.log(OPTION.EXPORT_ANIM + ":\t\t" + str(CONFIG.export_animation),
                 LOG.ADDINFO)
        self.log(OPTION.WRITE_LOG + ":\t\t\t" + str(CONFIG.write_log),
                 LOG.ADDINFO)
        self.log(OPTION.ONLY_XML + ":\t\t\t" + str(CONFIG.only_xml),
                 LOG.ADDINFO)
        self.log(OPTION.ERROR_CORRECTION + ":\t\t" +
                 str(CONFIG.error_correction).capitalize(),
                 LOG.ADDINFO)
#        self.log(OPTION.EXPORT_SCALED + ":\t\t" +
#                 str(CONFIG.export_scaled),
#                 LOG.ADDINFO)
#        self.log(OPTION.SCALING_FACTOR + ":\t\t" +
#                 str(round(CONFIG.scaling_factor, FPM.NDIGITS)),
#                 LOG.ADDINFO)
        self.log(OPTION.EXPORT_DIFFUSE_AS_AMBIENT + ":\t" +
                 str(CONFIG.export_diffuse_as_ambient),
                 LOG.ADDINFO)
        self.log(OPTION.EXPORT_MIRROR_AS_EMIT + ":\t\t" +
                 str(CONFIG.export_mirror_as_emit),
                 LOG.ADDINFO)
        self.log("Unit system:\t\t" +
                 str(self.context.scene.unit_settings.system).capitalize(),
                 LOG.ADDINFO)
        self.log("Unit scale:\t\t\t" +
                 str(self.context.scene.unit_settings.scale_length),
                 LOG.ADDINFO)
        self.log("", LOG.ADDINFO)
        ## collect & check bones
        self.console_message('collect and check exportable data')
        self.get_bones()
        self.get_and_check_root_bone()
        ## collect the rest
        self.get_meshes()
        self.get_materials()
        self.get_attachment_points()
        self.get_animation_basics()
        if CONFIG.export_animation:
            self.get_animation_events()
        ## check the rest
        self.check_meshes()
        self.check_materials()
        self.check_hierarchy()
        self.check_influence()
        self.check_names()
        self.console_message('data collected and checked')
        return self.status

    def write_data(self):
        '''write all data into the export file'''

        ## open output file
        self.console_message("create and write xml data")
        f = open(self.export_filename, mode="w", encoding="utf-8")
        ## write intro
        f.write("<trainzImport>\n" + IND1 + "<version>1</version>\n")
        ## write mesh section
        self.console_message("write mesh section")
        self.write_mesh_section(f)
        ## write skeleton section
        self.console_message("write skeleton section")
        self.write_skeleton_section(f)
        ## write material section
        self.console_message("write material section")
        self.write_material_section(f)
        ## write animation section
        if CONFIG.export_animation:
            self.console_message("write animation section")
            self.write_animation_section(f)
        ## write outro
        f.write("</trainzImport>\n")
        f.close()
        self.console_message("xml data written")

        ## invoke TMI if requested
        if not CONFIG.only_xml:
            self.console_message("hand over to TrainzMeshImporter:\n")
            if (os.name != 'nt'):
                self.log(
                    "TrainzMeshImporter.exe can only run on Windows systems. "
                    "Finish after writing XML file"
                    "(" + self.export_filename + ").",
                    LOG.ERROR)
            elif not os.path.exists(SCRIPT.PATH + TMI.FILENAME):
                self.log(
                    "file \"" + SCRIPT.PATH + TMI.FILENAME + "\" not found",
                    LOG.ERROR)
            else:
                cmd_line = []
                cmd_line.append(SCRIPT.PATH + TMI.FILENAME)
                cmd_line.append("-inFile")
                cmd_line.append(self.export_filename)
                cmd_line.append("-outFile")
                cmd_line.append(self.export_filename)
                cmd_line.append("-outputIM")
                cmd_line.append(str(bool(CONFIG.export_mesh)).lower())
                cmd_line.append("-outputKIN")
                cmd_line.append(str(bool(CONFIG.export_animation)).lower())
                if CONFIG.write_log:
                    cmd_line.append("-log")
                    cmd_line.append(self.TMI_log_filename)
                self.log("calling TMI:\t\"" + "\" \"".join(cmd_line) + "\"\n",
                         LOG.INFO)
                ret_code = subprocess.call(cmd_line)
                ## merge TMI log into BET log
                if CONFIG.write_log:
                    ## read TMI log
                    log_file = open(self.TMI_log_filename)
                    content = log_file.readlines()
                    log_file.close()
                    ## write into BET log
                    log_file = open(self.log_filename, "a")
                    log_file.write("\n")
                    log_file.writelines(content)
                    log_file.write("\n")
                    log_file.close()
                    # drop TMI log
                    while os.path.exists(self.TMI_log_filename):
                        try:
                            os.remove(self.TMI_log_filename)
                        except:
                            pass
                if ret_code != 0:
                    self.log("TrainzMeshImporter finished with error(s).",
                             LOG.ERROR)
        return self.status

    def export(self):
        ''''''
        # start message
        print('\n')  # to structure console output
        self.console_message("----- Exporter for Trainz - start -----", 1)
        # save the current blender states
        self.save_state()
        # collect and evaluate data
        self.collect_data()
        # write export files if all is OK
        if self.status == STATUS.ERROR:
            self.log("Error(s) during data collection, export aborted.",
                     LOG.INFO)
        else:
            self.write_data()
        # restore former project status
        self.restore_state()
        # end message & time
        duration = time.time() - self.start_time
        self.console_message(
            "===== Exporter for Trainz =  end  ===== "
            "runtime %(m)d:%(s)02d:%(ms)03d\n" % {
                'm': (int(duration) / 60),
                's': (int(duration) % 60),
                'ms': ((duration - int(duration)) * 1000)},
            1)


#### wrapper between operator and export class ##########

def start(filename, context):
    '''wrapper, to split the Blender GUI part from the Exporter part'''
    # create TrainzExport
    te = TrainzExport(filename, context)
    # start TrainzExport
    te.export()
    # return the result
    return te.status


#### user interface ###################################

class export_trainz(bpy.types.Operator):
    '''Export objects as XML/IM/KIN file to import into Trainz'''
    bl_idname = "export.trainz"
    bl_description = "Export objects as XML/IM/KIN file to import into Trainz"
    bl_label = "Export Trainz"

    filename_ext = ".xml"
    filter_glob = bpy.props.StringProperty(default="*.xml", options={'HIDDEN'})
    #list of operator properties
    filepath = (
        bpy.props.StringProperty(
            name="File Path",
            description="File path used for exporting the XML file.",
            maxlen=1024,
            default=""))
    check_existing = (
        bpy.props.BoolProperty(
            name="Check Existing",
            description="Check and warn on overwriting existing files.",
            default=True,
            options={'HIDDEN'}))
    # create UI properties
    selection_method = (
        bpy.props.EnumProperty(
            name="Selection Method",
            items=((SELECTIONMETHOD.VISIBLE,
                    'visible',
                    "export visible objects"),
                   (SELECTIONMETHOD.SELECTED,
                    'selected',
                    "export selected objects")),
            description=("Script recognizes and export either visible or "
                         "selected objects.")))
    export_mesh = (
        bpy.props.BoolProperty(
            name="Export Mesh Data",
            description="Create .im file from XML."))
    export_anim = (
        bpy.props.BoolProperty(
            name="Export Animation Data",
            description="Create .kin file from XML."))
    write_log = (
        bpy.props.BoolProperty(
            name="Write Log",
            description="Write export messages into log file."))
    only_xml = (
        bpy.props.BoolProperty(
            name="only XML",
            description="Export only to XML file and don't build binaries."))
    error_handling = (
        bpy.props.EnumProperty(
            name="Error handling",
            items=((ERRORHANDLING.CORRECT,
                    'correct',
                    "try to correct the problem"),
                   (ERRORHANDLING.COLLECT,
                    'collect',
                    "collect faulty elements for further handling"),
                   (ERRORHANDLING.NONE,
                    'nothing',
                    "only give out an error message")),
            description="What should happen if correctable errors occure?"))

#    export_scaled = bpy.props.BoolProperty( \
#        name="Scale exported Data", \
#        description="Apply scaling factor during export.")
#    scaling_factor = bpy.props.FloatProperty( \
#        name="Scaling Factor", \
#        description="Factor to scale your data if scaling is enabled.", \
#        min=0.001, \
#        max=100.0, \
#        soft_min=0.001, \
#        soft_max=100.0, \
#        precision=4)

    export_diffuse_as_ambient = (
        bpy.props.BoolProperty(name="Export Diffuse as Ambient",
                               description=("Export diffuse color also as "
                                            "ambient color.")))
    export_mirror_as_emit = (
        bpy.props.BoolProperty(name="Export Mirror as Emit",
                               description=("Export mirror color as "
                                            "emit color.")))
    save_config = (
        bpy.props.BoolProperty(name="save current configuration",
                               description=("make the current configuration "
                                            "to the default configuration"),
                               default=False))

    # this exporter don't need an active object, so we always return true
    def invoke(self, context, event):
        #print('invoke')
        # if no config file exists we create one using the default values
        try:
            f = open(SCRIPT.PATH + CONFIG.FILENAME, "r")
            f.close()
        except IOError:
            f = open(SCRIPT.PATH + CONFIG.FILENAME, "w")
            CONFIGFILE.Parser.write(f)
            f.close()
        # read & present default configuration
        CONFIGFILE.Parser.read(SCRIPT.PATH + CONFIG.FILENAME)
        self.properties.export_mesh = (
            CONFIGFILE.Parser.getboolean(CONFIGFILE.SECTION,
                                         OPTION.EXPORT_MESH))
        self.properties.export_anim = (
            CONFIGFILE.Parser.getboolean(CONFIGFILE.SECTION,
                                         OPTION.EXPORT_ANIM))
#        self.properties.export_scaled = (
#            CONFIGFILE.Parser.getboolean(CONFIGFILE.SECTION,
#                                         OPTION.EXPORT_SCALED)
        self.properties.export_diffuse_as_ambient = (
            CONFIGFILE.Parser.getboolean(CONFIGFILE.SECTION,
                                         OPTION.EXPORT_DIFFUSE_AS_AMBIENT))
        self.properties.export_mirror_as_emit = (
            CONFIGFILE.Parser.getboolean(CONFIGFILE.SECTION,
                                         OPTION.EXPORT_MIRROR_AS_EMIT))
#        self.properties.scaling_factor = (
#            CONFIGFILE.Parser.getfloat(CONFIGFILE.SECTION,
#                                       OPTION.SCALING_FACTOR))
        self.properties.write_log = (
            CONFIGFILE.Parser.getboolean(CONFIGFILE.SECTION,
                                         OPTION.WRITE_LOG))
        self.properties.only_xml = (
            CONFIGFILE.Parser.getboolean(CONFIGFILE.SECTION,
                                         OPTION.ONLY_XML))
        self.properties.error_handling = (
            CONFIGFILE.Parser.get(CONFIGFILE.SECTION,
                                  OPTION.ERROR_CORRECTION))
        self.properties.selection_method = (
            CONFIGFILE.Parser.get(CONFIGFILE.SECTION,
                                  OPTION.SELECTION_METHOD))
        #set default path
        if bpy.data.filepath == '':
            ## default the filepath to "my documents" like blender would do if
            ## no path is given, taken from "winpaths.py" made by Ryan Ginstrom
            import ctypes
            from ctypes import windll, wintypes
            try:
                _SHGetFolderPath = windll.shell32.SHGetFolderPathW
                _SHGetFolderPath.argtypes = [wintypes.HWND,
                                             ctypes.c_int,
                                             wintypes.HANDLE,
                                             wintypes.DWORD,
                                             wintypes.LPCWSTR]
                path_buf = wintypes.create_unicode_buffer(wintypes.MAX_PATH)
                result = _SHGetFolderPath(0, 5, 0, 0, path_buf)
                self.properties.filepath = path_buf.value + "\\untitled.xml"
            except Exception:
                self.properties.filepath = "untitled.xml"
        else:
            self.properties.filepath = bpy.data.filepath.replace(".blend",
                                                                 ".xml")
        # orig code
        context.window_manager.fileselect_add(self)
        return {'RUNNING_MODAL'}

    def execute(self, context):
        #print('execute')
        # save user input
        CONFIG.export_mesh = self.properties.export_mesh
        CONFIG.export_animation = self.properties.export_anim
#        CONFIG.export_scaled = self.properties.export_scaled
        CONFIG.export_diffuse_as_ambient = (
            self.properties.export_diffuse_as_ambient)
        CONFIG.export_mirror_as_emit = self.properties.export_mirror_as_emit
#        CONFIG.scaling_factor = self.properties.scaling_factor
        CONFIG.write_log = self.properties.write_log
        CONFIG.only_xml = self.properties.only_xml
        CONFIG.error_correction = self.properties.error_handling
        CONFIG.selection_method = self.properties.selection_method
        # save config if requested
        if self.properties.save_config:
            # update config file parser
            CONFIGFILE.Parser.set(CONFIGFILE.SECTION,
                                  OPTION.EXPORT_MESH,
                                  str(CONFIG.export_mesh))
            CONFIGFILE.Parser.set(CONFIGFILE.SECTION,
                                  OPTION.EXPORT_ANIM,
                                  str(CONFIG.export_animation))
#            CONFIGFILE.Parser.set(CONFIGFILE.SECTION,
#                                  OPTION.EXPORT_SCALED,
#                                  str(CONFIG.export_scaled))
            CONFIGFILE.Parser.set(CONFIGFILE.SECTION,
                                  OPTION.EXPORT_DIFFUSE_AS_AMBIENT,
                                  str(CONFIG.export_diffuse_as_ambient))
            CONFIGFILE.Parser.set(CONFIGFILE.SECTION,
                                  OPTION.EXPORT_MIRROR_AS_EMIT,
                                  str(CONFIG.export_mirror_as_emit))
#            CONFIGFILE.Parser.set(CONFIGFILE.SECTION,
#                                  OPTION.SCALING_FACTOR,
#                                  str(round(CONFIG.scaling_factor,
#                                            FPM.NDIGITS)))
            CONFIGFILE.Parser.set(CONFIGFILE.SECTION,
                                  OPTION.WRITE_LOG,
                                  str(CONFIG.write_log))
            CONFIGFILE.Parser.set(CONFIGFILE.SECTION,
                                  OPTION.ONLY_XML,
                                  str(CONFIG.only_xml))
            CONFIGFILE.Parser.set(CONFIGFILE.SECTION,
                                  OPTION.ERROR_CORRECTION,
                                  str(CONFIG.error_correction))
            CONFIGFILE.Parser.set(CONFIGFILE.SECTION,
                                  OPTION.SELECTION_METHOD,
                                  str(CONFIG.selection_method))
            # rewrite config file
            with open(SCRIPT.PATH + CONFIG.FILENAME, "w") as f:
                CONFIGFILE.Parser.write(f)
            f.close()
        # call exporter
        result = start(self.properties.filepath, context)
        ## inform about the outcome
        if result == 0:  # -> OK
            self.report({'INFO'},
                        "Export finished successfully.")
        elif result == 1:  # -> WARNINGS
            self.report({'WARNING'},
                        ("Export finished with warnings. "
                         "Please check Blender console window."))
        else:  # -> ERRORS
            self.report({'ERROR'},
                        "Export failed. Please check Blender console window.")
        return {'FINISHED'}


## register script as operator and add it to the File->Export menu

def menu_func(self, context):
    self.layout.operator(export_trainz.bl_idname,
                         text="Trainz Mesh and Animation...")


def register():
    bpy.utils.register_module(__name__)
    bpy.types.INFO_MT_file_export.append(menu_func)


def unregister():
    bpy.utils.unregister_module(__name__)
    bpy.types.INFO_MT_file_export.remove(menu_func)

if __name__ == "__main__":
    register()

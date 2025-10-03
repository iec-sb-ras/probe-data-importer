"""
Модуль для обработки геохимических данных из Excel-файлов и преобразования в RDF-формат. Также можно загружать результат в хранилище Графа знаний.

Основные компоненты:
- Импорт данных из XLS-файлов с помощью xlrd
- Преобразование химических и геологических данных в RDF-триплы
- Генерация онтологических связей между образцами, измерениями и локациями
- Загрузка результата в удалённое хранилище через HTTP PUT

Основные классы:
- ImpState: Базовый класс для обработки данных Excel-листа
- Yarki, Kharantsy, Khuzhir: Конкретные реализации для различных форматов данных

Основные функции:
- parse_xl(): Парсинг Excel-файла с использованием указанного обработчика
- upload(): Загрузка сгенерированного RDF-файла на сервер
- normURI(): Нормализация строк для URI
- elem(): Получение IRI элемента по символу

Константы:
- FILES: Словарь соответствия файлов обработчикам и локациям
- REGEX-паттерны для обработки химических формул и координат
- Пространства имен RDF (PT, P, MT, WGS и др.)

Использование:
1. Определите обработчик данных в FILES
2. Вызовите parse_xl() для обработки файла
3. Используйте upload() для отправки результатов

Для замены реализации:
1. Создайте новый класс, наследующий от ImpState (или его подклассов)
2. Реализуйте методы обработки данных в соответствии с новой структурой файлов
3. Обновите словарь FILES для соответствия новым файлам и обработчикам
4. При необходимости адаптируйте регулярные выражения и логику преобразования

Пример использования:
    parse_xl("filename.xls", (Yarki, ["Локация"]))
    upload("output.ttl")

"""

# import pandas as pd
import xlrd
from rdflib import (Graph, Namespace, FOAF, XSD, RDF, RDFS, DCTERMS, URIRef,
                    Literal, BNode)
from rdflib.namespace import WGS, SDO
import os.path
import unicodedata
from enum import Enum
import re
import requests as rq
from requests.auth import HTTPBasicAuth
import base64
import os
from namespace import PT, P, SCHEMA, BIBO, MT, GS, CGI, DBP, DBP_OWL
from pprint import pprint

import pudb

try:
    del os.environ["HTTP_PROXY"]
    del os.environ["HTTPS_PROXY"]
except KeyError:
    pass
try:
    del os.environ["http_proxy"]
    del os.environ["https_proxy"]
except KeyError:
    pass

ONTODIR = "../data/kg"
SUBDIR = "../data/xenolites/"

TARGET = "alrosa.ttl"
TARGETMT = "PeriodicTable.ttl"

EQUIPMENT = "S8 Tiger"
EQUIPMENT_TYPE = "X-Ray fluorescence analysis"

G = Graph(bind_namespaces="rdflib")
GMT = Graph(bind_namespaces="rdflib")
try:
    GMT.parse(location=os.path.join(ONTODIR, 'PeriodicTable.owl'))
except FileNotFoundError:
    pass
GS = [G, GMT]

COMPRE = re.compile(r"^(([A-Z][A-Za-z]{,2}\d{,2})+)(.*?)$")
COMPELRE = re.compile(r"[A-Z][A-Za-z]{,2}\d{,2}")
# r'([A-Z][a-z]*)(\d+(?:\.\d+)?)?'
ELRE = re.compile(r"^([A-Z][A-Za-z]{,2})+(.*)$")

DEGRE = re.compile(r"^(\d{,3})(\D)(\d{,2})(\D)(\d{,2}(\.\d{,2})?)(\D)(.*)$")
# 107°25'28.48"В
# m = DEGRE.match('107°25\'28.48"В')
# print(m, m.groups())
# m = DEGRE.match('107°25\'28"В')
# print(m, m.groups())
# quit()

# spl = COMPELRE.findall("C2H5W8K9OH")
# print(spl)
# spl = COMPELRE.findall("OH")
# print(spl)
# spl = COMPELRE.findall("SiO2")
# print(spl)
# # quit()

COMPOUNDS = {}

for _ in GS:
    _.bind("pt", PT)
    _.bind("pi", P)
    _.bind("mt", MT)
    _.bind("bibo", BIBO)
    _.bind("gs", GS)
    _.bind("cgi", CGI)
    _.bind("dbp", DBP)
    _.bind("dbp_owl", DBP_OWL)
    _.bind("wgs", WGS)
    # _.bind('geo', GEO)

ElToIRI = {}
GeoSite = PT.Site
DataSheet = PT.DataSheet
GeoSample = PT.Sample
GeoMeasure = PT.Measurement
samplerel = PT.sample
PPM = PT.PPM
Percent = PT.Percent
SpatialThing = WGS.SpatialThing
Long = WGS.long
Lat = WGS.lat
IgnitionLosses = PT.IgnitionLosses

G.add((PT.PPM, RDFS.label, Literal("мг/кг", lang="ru")))
G.add((PT.Percent, RDFS.label, Literal("Процент", lang="ru")))

for el in GMT.subjects(RDF.type, MT.Element):
    ElToIRI[el.fragment] = el

# print(ElToIRI)


def load_lithology_ontology(graph=None,
                            ttl_file="lithology.ttl",
                            ontodir="./"):
    """
    Загружает онтологию литологии из TTL-файла в RDF-граф.

    Args:
        graph (Graph): RDF-граф для загрузки онтологии. Если None, создается новый.
        ttl_file (str): Имя TTL-файла с онтологией литологии.
        ontodir (str): Директория с онтологическими файлами.

    Returns:
        Graph: Граф с загруженной онтологией литологии.
    """
    if graph is None:
        graph = Graph(bind_namespaces="rdflib")

    # Определяем полный путь к файлу
    ttl_path = os.path.join(ontodir, ttl_file)

    try:
        # Загружаем онтологию литологии
        graph.parse(location=ttl_path, format="turtle")
        print(f"✓ Онтология литологии загружена из: {ttl_path}")
        print(f"✓ Триплетов в графе после загрузки: {len(graph)}")

        # Проверяем основные классы онтологии
        check_ontology_classes(graph)

    except FileNotFoundError:
        print(f"⚠ Файл онтологии не найден: {ttl_path}")
        print("  Создайте файл lithology.ttl с онтологией литологии")
    except Exception as e:
        print(f"✗ Ошибка загрузки онтологии: {e}")

    return graph


load_lithology_ontology(G)

def snake_to_camel(snake_str: str, capitalize_first: bool = False) -> str:
    """
    Convert snake_case string to camelCase.

    Args:
        snake_str: The snake_case string to convert
        capitalize_first: If True, returns UpperCamelCase (PascalCase),
                         if False, returns lowerCamelCase (default)

    Returns:
        The converted camelCase string
    """
    if not snake_str:
        return snake_str

    # Split by underscore and capitalize each word except first one
    parts = snake_str.split('_')

    if capitalize_first:
        # UpperCamelCase (PascalCase) - capitalize all words
        return ''.join(part.capitalize() for part in parts)
    else:
        # lowerCamelCase - capitalize all words except first
        return parts[0] + ''.join(part.capitalize() for part in parts[1:])

def capitalize(words):
    orig = " ".join(words.lower().split())
    orig = orig.capitalize()
    return orig

def check_ontology_classes(graph):
    """
    Проверяет наличие основных классов онтологии литологии в графе.

    Args:
        graph (Graph): RDF-граф для проверки.
    """
    # Определяем namespace для литологии (предполагаем, что он используется в TTL)
    LITHO = Namespace("http://example.org/geology/")
    graph.bind("litho", LITHO)

    # Проверяем основные классы
    core_classes = {
        "RockType": "Тип горной породы",
        "Mineral": "Минерал",
        "GeologicStructure": "Геологическая структура",
        "UltramaficRock": "Ультрамафическая порода",
        "MetamorphicRock": "Метаморфическая порода"
    }

    print("\n=== ПРОВЕРКА КЛАССОВ ОНТОЛОГИИ ===")

    for class_name, description in core_classes.items():
        # Проверяем в разных namespace'ах
        class_iris = [
            LITHO[class_name],
            URIRef(f"http://example.org/geology/{class_name}"),
            URIRef(
                f"http://crust.irk.ru/ontology/pollution/terms/1.0/{class_name}"
            )
        ]

        found = False
        for class_iri in class_iris:
            if (class_iri, RDF.type, RDFS.Class) in graph:
                print(f"✓ Найден класс: {class_name} ({description})")
                found = True
                break

        if not found:
            print(f"⚠ Класс не найден: {class_name}")


def normURI(s):
    """Нормализует строку для использования в URI.

    Оставляет только буквы и цифры, заменяет остальные символы на подчеркивания.

    Args:
        s (str): Входная строка.

    Returns:
        str: Нормализованная строка.
    """

    r = ""
    for c in s:
        d = unicodedata.category(c)
        if d[0] in ["L", "N"]:
            r += c
        elif len(r) > 0 and r[-1] != "_":
            r += "_"
    return r.rstrip("_")


def elem(name):
    """Возвращает IRI элемента по его символу.

    Args:
        name (str): Символ элемента (например, 'Fe').

    Returns:
        URIRef: IRI элемента, или None, если не найден.
    """

    p1, p2 = name[:1], name[1:]
    p1 = p1.upper()
    p2 = p2.lower()
    return ElToIRI.get(p1 + p2, None)


PPM_TO_PERCENT = 0.0001  # 1 PPM = 0.0001%

def convert_units(value, from_unit, to_unit='%'):
    """
    Конвертирует значения между единицами измерения
    """
    if from_unit == to_unit:
        return value

    conversion_factors = {
        ('PPM', '%'): PPM_TO_PERCENT,
        ('%', 'PPM'): 1 / PPM_TO_PERCENT,
    }

    factor = conversion_factors.get((from_unit, to_unit))
    if factor:
        return value * factor
    else:
        return value  # Если конвертация невозможна, возвращаем исходное значение

# Словарь для отображения текстур на IRI онтологии
ROCK_TEXTURE_MAPPING = {
    # Основные текстуры
    'COARSE': PT.CoarseGrainedTexture,
    'COARSE-GRAINED': PT.CoarseGrainedTexture,
    'MEDIUM-GRAINED': PT.MediumGrainedTexture,
    'FINE-GRAINED': PT.FineGrainedTexture,
    'FINE': PT.FineGrainedTexture,

    # Породные текстуры
    'GRANOBLASTIC': PT.GranoblasticTexture,
    'PORPHYROCLASTIC': PT.PorphyroclasticTexture,
    'PORPHYRITIC': PT.PorphyriticTexture,
    'CUMULATE': PT.CumulateTexture,
    'MEGACRYSTALLINE': PT.MegacrystallineTexture,
    'MEGACRYSTIC': PT.MegacrystallineTexture,
    'GRANULAR': PT.GranularTexture,
    'MOSAIC': PT.MosaicTexture,
    'EQUIGRANULAR': PT.EquigranularTexture,
    'SHEARED': PT.ShearedTexture,
    'FOLIATED': PT.FoliatedTexture,
    'FLUIDAL': PT.FluidalTexture,
    'BANDED': PT.BandedTexture,
    'LAMINAR': PT.LaminatedTexture,
    'LAMINATED': PT.LaminatedTexture,

    # Комбинированные текстуры
    'COARSE PORPHYRITIC': PT.CoarsePorphyriticTexture,
    'FINE PORPHYRITIC': PT.FinePorphyriticTexture,
    'COARSE-PORPHYRIC': PT.CoarsePorphyriticTexture,
    'COARSE PORPHYROCLASTIC': PT.CoarsePorphyroclasticTexture,
    'FINE-GRAINED PORPHYROCLASTIC': PT.FineGrainedPorphyroclasticTexture,
    'FINE PORPHYROCLASTIC': PT.FineGrainedPorphyroclasticTexture,
    'MOSAIC-PORPHYROCLASTIC': PT.MosaicPorphyroclasticTexture,
    'GRANOBLASTIC, BANDED': PT.GranoblasticBandedTexture,
    'COARSE GRANOBLASTIC': PT.CoarseGranoblasticTexture,
    'COARSE EQUANT': PT.CoarseEquantTexture,
    'COARSE-EQUANT': PT.CoarseEquantTexture,
    'COARSE TABULAR': PT.CoarseTabularTexture,
    'COARSE-TABULAR': PT.CoarseTabularTexture,
    'MEDIUM- TO COARSE-GRAINED': PT.MediumToCoarseGrainedTexture,
    'FINE- TO MEDIUM-GRAINED': PT.FineToMediumGrainedTexture,
    'FINE- TO COARSE-GRAINED': PT.FineToMediumGrainedTexture,
    'HETEROGRANULAR': PT.HeterogeneousTexture,
    'ALLOTRIOMORPHIC GRANULAR': PT.AllotriomorphicGranularTexture,
    'XENOMORPHIC GRANULAR': PT.XenomorphicGranularTexture,
    'INTERLOCKING': PT.InterlockingTexture,
    'DEFORMED': PT.DeformedTexture,
    'TRANSITIONAL': PT.TransitionalTexture,
    'MYLONITIC': PT.MyloniticTexture,
    'FLUIDAL MOSAIC': PT.FluidalMosaicTexture,

    # Специфичные текстуры
    'COARSE LAMELLAR': PT.CoarseLamellarTexture,
    'COARSE PROTOGRANULAR': PT.CoarseProtogranularTexture,
    'PROTOGRANULAR': PT.ProtogranularTexture,

    # Сложные комбинированные текстуры
    'LAYERED AND DISRUPTED MOSAIC PORPHYROCLASTIC': PT.LayeredDisruptedMosaicPorphyroclasticTexture,
    'FLUIDAL MOSAIC PORPHYROCLASTIC': PT.FluidalMosaicTexture,
    'MOSAIC-TABULAR TO EQUIGRANULAR': PT.MosaicTexture,
    'HYPAUTOMORPHIC, OPHITIC WITH TRANSITION TO GRANO-LAPIDOBLASTIC': PT.HypautomorphicOphiticTexture,
    'FASCICULATE': PT.FasciculateTexture,
    'POLYGONAL GRANOBLASTIC': PT.PolygonalGranoblasticTexture,

    # Порфиробластические текстуры
    'PORPHYROBLASTIC': PT.PorphyroclasticTexture,

    # Упрощенные маппинги для сложных комбинаций
    'COARSE-PORPHYRIC, PORPHYROCLASTIC, WEAKLY LAMINAR': PT.ComplexTexture,
    'COARSE GRANULAR': PT.CoarseGrainedTexture,
    'COARSE-GRANULAR': PT.CoarseGrainedTexture,
    'MEDIUM-EQUANT': PT.EquigranularTexture,
    'MEDIUM-TABULAR': PT.MediumGrainedTexture,
    'COARSE-EQUANT/TABULAR FOLIATED': PT.CoarseEquantTexture,
    'COARSE-TABULAR FOLIATED': PT.CoarseTabularTexture,

    # Обработка опечаток
    'COASRSE': PT.CoarseGrainedTexture,
    'GRANULOBLASTIC': PT.GranoblasticTexture
}

def get_texture_iri(texture_string):
    """
    Преобразует строку с описанием текстуры в IRI онтологии.

    Args:
        texture_string (str): Строка с описанием текстуры горной породы

    Returns:
        URIRef: IRI соответствующего класса текстуры или PT.ComplexTexture если не найдено
    """
    if not texture_string:
        return None

    # Нормализация строки: верхний регистр и удаление лишних пробелов
    normalized = ' '.join(texture_string.upper().split())

    # Прямое соответствие
    if normalized in ROCK_TEXTURE_MAPPING:
        return ROCK_TEXTURE_MAPPING[normalized]

    # Попытка найти частичное соответствие для сложных описаний
    for key, value in ROCK_TEXTURE_MAPPING.items():
        if key in normalized:
            return value

    # Возвращаем общий класс для неизвестных текстур
    return PT.ComplexTexture


class State(Enum):
    NONE = 0
    CLASS = 1
    HEADER = 2
    DATA = 3
    IGNORE = 4
    LOCATION = 5
    DETLIM = 6
    REFERENCES = 7


class ImpState:
    """
    Базовый класс для обработки данных из Excel-листа.

    Атрибуты:
    - state: Текущее состояние обработки
    - g: RDF-граф для записи результатов
    - locations: Список обработанных локаций

    Методы:
    - proc_comp(): Обработка химических соединений и элементов
    - proc_loc(): Добавление новой локации
    - r(): Обработка строки данных
    - c(): Обработка отдельной ячейки

    Базовый класс для обработки листа Excel.

    Attributes:
        state (State): Текущее состояние обработки.
        instr (bool): Флаг, указывающий, что текущая строка - инструкция.
        prev_class: Предыдущий класс.
        cls (dict): Словарь классов для колонок.
        header (dict): Словарь заголовков для колонок.
        dsiri: IRI набора данных.
        g (Graph): RDF-граф.
        sample: Текущий образец.
        locations (list): Список локаций.
        sample_col (int): Номер колонки с образцами.
        loc_fence (int): Граница для локаций.
        dlims (dict): Словарь пределов обнаружения.

    Methods:
        proc_comp: Обрабатывает химическое соединение или элемент.
        proc_locs: Обрабатывает список локаций.
        proc_loc: Обрабатывает одну локацию.
        belongs: Устанавливает отношение принадлежности.
        data: Проверяет, находится ли в состоянии DATA.
        ign: Проверяет, находится ли в состоянии IGNORE.
        r: Обрабатывает строку.
        c: Обрабатывает ячейку в режиме данных.
        h: Обрабатывает ячейку заголовка.
        hc: Обрабатывает ячейку класса.
    """

    _sample_iri_ = samplerel
    _belongs_iri_ = PT["location"]
    _start_state_ = None

    def __init__(self, graph, datasetIRI, locations=None, **kwargs):
        self.state = self._start_state_
        self.instr = False
        self.prev_class = None
        self.cls = {}
        self.header = {}
        self.dsiri = datasetIRI  # TODO: Add type GelologicalSite!!
        self.g = graph
        self.sample = None
        self.locations = []
        self.sample_col = None
        self.proc_locs(locations)
        self.loc_fence = len(self.locations)
        self.dlims = {}
        self.kwargs = kwargs
        self.non_iso = {}
        self.measurements = {}
        self.analysis = None
        self.fls = {}

    def proc_value(self, value):
        if isinstance(value, str):
            sv = value.strip()
            return sv
        sn = value
        if isinstance(sn, float) and sn.is_integer():
            return int(sn)
        return sn

    def add(self, triple):
        # print("T:{}".format(triple))
        self.g.add(triple)

    def proc_comp(self, names, value, delim=False):
        """
        Обрабатывает химическое соединение или элемент с 	сохранением оригинальных значений
        и добавлением нормализованных (PPM -> %)

        Аргументы:
        - names: Кортеж (поле, имя_поля)
        - value: Значение ячейки
        - delim: Флаг предела обнаружения
                """

        uvalue = str(value).upper().strip()
        if isinstance(value, str):
            value = value.strip()
        if "N.A." in uvalue:
            return
        dl = None
        ovalue = value  # Original value
        if "<" in uvalue:
            dl = value.lstrip("<").strip()
            value = dl

        name, fieldname = names
        name = name.strip()
        mo = COMPRE.match(name)
        add = self.add

        rel = PT[normURI(name)]

        def finish():
            if self.analysis and not delim:
                # print(type(ovalue))
                ### print("->{}->{}".format(rel, repr(ovalue)))
                add((self.analysis, rel, Literal(ovalue)))

        def degs(v):
            if isinstance(v, float):
                return v
            # 107°25'28.48"В
            m = DEGRE.match(v)
            if m is None:
                return v
            d, _, m, _, s, _, _, dir = m.groups()
            d, m, s = [float(v) for v in [d, m, s]]
            d += m / 60.0
            d += s / 3600.0
            return d

        def unit(m, rupper):
            if "PPM" in rupper:
                add((m, PT.unit, PPM))
            elif "INT" in rupper:
                add((m, PT.unit, P.Int))
            elif "%" in fieldname:
                add((m, PT.unit, Percent))
            else:
                add((m, PT.unit, P.UnknowUnit))

        def create_measurement_with_normalization(m, measurement_value, unit_type, rupper):
            """Создает измерение с нормализованными значениями"""
            # Сохраняем оригинальное значение
            if dl is None:
                add((m, PT.value, Literal(measurement_value)))

            # Определяем единицы измерения
            if unit_type == 'PPM':
                add((m, PT.unit, PPM))
                # 🔥 ДОБАВЛЯЕМ НОРМАЛИЗОВАННОЕ ЗНАЧЕНИЕ (PPM -> %)
                if not delim and isinstance(measurement_value, (int, float)):
                    normalized_value = measurement_value * 0.0001  # PPM to %
                    add((m, PT.normalizedValue, Literal(normalized_value)))
                    add((m, PT.normalizedUnit, Percent))
            elif unit_type == '%':
                add((m, PT.unit, Percent))
            elif unit_type == 'INT':
                add((m, PT.unit, P.Int))
            else:
                add((m, PT.unit, P.UnknowUnit))

        if isinstance(value, float) and value.is_integer():
            value = int(value)
        if name in ["ППП", "ппп"]:
            rel = PT.il  # ignition losses
            m = BNode()
            add((self.analysis, PT.measurement, m))
            add((m, PT.value, Literal(value)))
            add((m, RDF.type, GeoMeasure))
            add((m, RDF.type, IgnitionLosses))
            u = name.upper()
            unit(m, u)
            return

        if mo is None:
            if name == "с_ш":
                rel = Lat
                ovalue = degs(ovalue)
            if name == "в_д":
                rel = Long
                ovalue = degs(ovalue)
            if name in ["sku", "номер"]:
                rel = SDO.sku
                if isinstance(ovalue, float):
                    ovalue = int(ovalue)
            finish()
            return
        else:
            comp = mo.group(1)
            rest = mo.group(3)
            rc = COMPELRE.findall(comp)
            el1 = rc[0]
            el = ELRE.match(el1).group(1)
            eliri = elem(el)

        if eliri is None:  # This is not a compound
            if self._field_map_ is not None:
                f = self._field_map_.get(fieldname, None)
                if f is None:
                    f = self._field_map_.get(name, None)
                if f is None:
                    finish()
                else:
                    f(self, value, names=names, delim=delim)
            return

        def make_detlim():
            m = BNode()
            add((self.dsiri, PT.detectionLimit, m))
            add((m, RDF.type, PT.DetectionLimit))
            return m

        m = None

        if self.analysis and not delim:
            m = BNode()
            add((self.analysis, PT.measurement, m))
            add((m, RDF.type, GeoMeasure))
        if delim:
            m = make_detlim()

        rupper = rest.upper()

        if dl is None:
            add((m, PT.value, Literal(value)))
            unit(m, rupper)

        # 🔥 СОЗДАЕМ ИЗМЕРЕНИЕ С НОРМАЛИЗАЦИЕЙ
        create_measurement_with_normalization(m, value, unit_type, rupper)

        def finish_dl(e, m):
            if delim:
                self.dlims[e] = m
            if dl is not None:
                dlm = self.dlims.get(e, None)
                if dlm is not None:
                    add((m, PT.value, dlm))
                else:
                    # m is description, connected to a sample
                    md = make_detlim()
                    # Для пределов обнаружения тоже добавляем нормализацию
                    add((md, PT.value, Literal(value)))
                    add((md, RDF.type, GeoMeasure))
                    unit(md, rupper)
                    create_measurement_with_normalization(md, value, unit_type, rupper)
                    self.dlims[e] = md
                    add((m, PT.value, md))

        if len(rc) > 1 or el1 != el:  # Compound, e.g. oxide
            compname = "compound-" + normURI(comp)
            cb = COMPOUNDS.get(compname, None)
            if cb is None:
                cb = PT[compname]
                add((cb, RDF.type, PT.Compound))
                add((cb, PT.Formula, Literal(comp)))
                add((cb, RDFS.label, Literal(comp)))
                COMPOUNDS[compname] = cb
            add((m, PT.compound, cb))

            finish_dl(comp, m)

        elif len(rc) == 1 and el1 == el and eliri is not None:
            add((m, MT.element, eliri))
            finish_dl(el, m)
        else:
            print("#!ERROR unknown combination of {}, and {}=?={}: {}.".format(
                rc, el1, el, eliri))
            quit()

        if "TOT" in rupper or "ОБЩ" in rupper:
            add((m, PT.total, Literal(True)))

    def proc_locs(self, locations):
        if locations is None:
            return
        for loc in locations:
            self.proc_loc(loc)

    def proc_loc(self, location):
        location = location.strip()
        uriloc = P[normURI(location)]
        locs = self.locations
        if locs:
            prev = locs[-1]
            self.belongs(uriloc, prev)
        add = self.add
        add((uriloc, RDFS.label, Literal(location, lang="ru")))
        add((uriloc, RDF.type, GeoSite))
        self.locations.append(uriloc)

    def belongs(self, obj, location=None):
        if location is None:
            if len(self.locations) > 0:
                location = self.locations[-1]
            else:
                return  # None to belong to
        add = self.add
        # print(self.locations)
        add((obj, self._belongs_iri_, location))

    def data(self):
        return self.state == State.DATA

    def ign(self):
        return self.state == State.IGNORE

    def row(self, row, rx):
        self.instr = False
        self.sample = None

        f = row[0]
        fs = f.value

        if f.ctype == xlrd.XL_CELL_TEXT:
            fss = fs.strip()
            if fss.startswith("#"):
                fss = fss.lstrip("#")
                try:
                    self.state = State[fss]
                except KeyError as k:
                    print("#! ERROR Key: {}".format(k))
                    quit()
                self.instr = True
        if self.ign() or self.instr:
            return

        detlim = self.state == State.DETLIM
        if self.data() or detlim:
            if self.sample is None:
                self.c(row[self.sample_col], rx, self.sample_col, sheet_row = row, detlim=delim)
            for i, cell in enumerate(row):
                self.c(cell, rx, i, detlim=delim, sheet_row=row)
            return

        if self.state == State.HEADER:
            for i, cell in enumerate(row):
                self.h(cell, rx, i)
            return

        if self.state == State.CLASS:
            self.prev_class = None
            for i, cell in enumerate(row):
                self.hc(cell, rx, i)
            return

        if self.state == State.LOCATION:
            for cell in row:
                if cell.ctype == xlrd.XL_CELL_TEXT:
                    if len(self.locations) > self.loc_fence:
                        self.locations.pop()
                    loc = cell.value.strip()
                    self.proc_loc(loc)

    def c(self, cell, row, col, sheet_row, detlim=False):
        if cell.ctype in [xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK]:
            return
        try:
            field, fieldname = self.header[col]
        except KeyError as k:
            print("#! ERROR header key {} not in {} row: {}".format(
                k, self.header, row))
            # quit()
            return

        prt = self.cls.get(col, "")
        if "%" in prt and "PPM" not in field:
            prt = "_%"
        elif "PPM" in prt and "%" not in field:
            prt = "_PPM"
        else:
            prt = ""

        add = self.add
        ds = self.dsiri
        val = cell.value
        if field == self._sample_iri_:
            if isinstance(val, str):
                name = val.replace(" ", "")
            else:
                if isinstance(val, float):
                    if val.is_integer():
                        name = str(int(val))
                    else:
                        name = "{}".format(val)
                else:
                    name = "{}".format(val)
            add((ds, self._sample_iri_, self.sample))
            add((self.sample, RDF.type, GeoSample))
            add((self.sample, RDFS.label, Literal(name)))
            # add((self.sample, RDF.type, SpatialThing))
            self.belongs(self.sample)
        elif self.sample is not None:
            self.proc_comp((field + prt, fieldname + prt), cell.value)
        elif detlim:
            self.proc_comp((field + prt, fieldname + prt), cell.value, detlim)
        else:
            print("#! ERROR: nowhere to store {} R:{} C:{}\n#!{}".format(
                cell, row, col, self.header))
            quit()

    def h(self, cell, row, col):
        if cell.ctype in [xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK]:
            return
        orig = str(cell.value)
        name = normURI(orig)
        if name in self._sample_names_:
            name = self._sample_iri_
            self.sample_col = col
        self.header[col] = (name, orig)

    def hc(self, cell, row, col):
        if cell.ctype in [xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK]:
            if self.prev_class is None:
                return
            orig, name = self.prev_class
        else:
            orig = str(cell.value).replace("мг/кг", "PPM")
            name = normURI(orig)
        self.cls[col] = orig
        self.prev_class = orig, name


class Yarki(ImpState):
    """Обработчик для данных формата Ярки."""
    _sample_names_ = ["Sample", "Полевой", "sample"]


class Kharantsy(ImpState):
    _sample_names_ = ["Номер_пробы"]


class Khuzhir(Kharantsy):
    pass


class Alrosa(ImpState):

    _start_state_ = State.HEADER
    _sample_names_ = ['SAMPLE_NAME']
    _sheet_names_ = [0]

    def reffield(self, value):
        try:
            cnum, reference = value.split(maxsplit=1)
        except ValueError:
            cnum = value.strip()
            reference = ""
        num = cnum.strip().lstrip("[").rstrip("]")
        refURI = P["ref-{}".format(num)]
        reference = reference.strip()
        if not reference:
            reference = None
        return refURI, reference

    def row(self, row, rx):
        self.instr = False
        self.sample = None
        self.analysis = None
        if self.state == State.HEADER:
            for i, cell in enumerate(row):
                self.h(cell, rx, i)
            self.state = State.DATA
            print("HEADER:{}".format(self.header))
            return
        c0 = row[0]
        v0 = str(c0.value).strip()
        if v0.startswith("#REFERENCES"):
            self.state = State.REFERENCES
            return
        if self.state == State.DATA:
            if self.sample is None and self.sample_col is not None:
                self.c(row[self.sample_col], rx, self.sample_col, sheet_row=row)
            for i, cell in enumerate(row):
                if i == self.sample_col:
                    continue
                self.c(cell, rx, i, sheet_row=row)
        if self.state == State.REFERENCES:
            refURI, reference = self.reffield(v0)
            assert (reference is not None)
            self.add((refURI, RDF.type, BIBO["AcademicArticle"]))
            self.add((refURI, RDFS.label, Literal(reference)))

    def c(self, cell, row, col, sheet_row, detlim=False):
        if cell.ctype in [xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK]:
            return
        try:
            field, fieldname = self.header[col]
        except KeyError as k:
            print("#! ERROR header key {} not in {} row: {}".format(
                k, self.header, row))
            # quit()
            return

        # prt = self.cls.get(col, "")
        prt = ""
        if "%" in prt and "PPM" not in field:
            prt = "_%"
        elif "PPM" in prt and "%" not in field:
            prt = "_PPM"
        else:
            prt = ""

        add = self.add
        ds = self.dsiri
        val = cell.value
        if field == self._sample_iri_:
            if isinstance(val, str):
                name = val.replace(" ", "")
            else:
                if isinstance(val, float):
                    if val.is_integer():
                        name = str(int(val))
                    else:
                        name = "{}".format(val)
                else:
                    name = "{}".format(val)
            name = name.strip().lstrip('samp.')
            name_orig = name
            name = name.replace("^A", "")
            name = name.replace("^D", "")
            name = name.replace("^M", "")
            name = name.replace("/", "-sl-")
            name = name.replace("?", "-q-")
            name = name.strip(" ")
            name = name.lstrip()
            if name_orig!=name:
                # print("PROBLEMATIC:{} ({})".format(name, name_orig))
                ns = self.non_iso.setdefault(name, [])
                ns.append(name_orig)
            samplename = 'sample-'+name
            sample_iri = P[samplename]
            meas = self.measurements
            mlist = meas.setdefault(sample_iri, [])
            self.analysis = P['analysis-{}-{}'.format(len(mlist)+1, samplename)]
            mlist.append(self.analysis)
            # if (sample_iri, RDF.type, PT.GeoSample) in self.g:
            #     print("Double: {} \n ROW:{}".format())
            #     quit
            self.sample = sample_iri
            add((ds, self._sample_iri_, self.sample))
            add((self.sample, RDF.type, GeoSample))
            add((self.sample, RDFS.label, Literal(name)))
            add((self.sample, PT.hasAnalysis, self.analysis))

            # add((self.sample, RDF.type, SpatialThing))
            self.belongs(self.sample)
        elif field == "CITATION":
            assert isinstance(val, str)
            refURI, reference = self.reffield(val)
            add((self.sample, BIBO.cites, refURI))
        elif field == "LOCATION":
            assert isinstance(val, str)
            locations = val.split("/")
            locations = [l.strip() for l in locations]

            for location in locations:
                if location not in self.fls:
                    location_bnode = BNode()  # Создаем BNode для каждой локации
                    self.fls[location] = location_bnode

                    # Добавляем основную информацию о локации
                    add((location_bnode, RDF.type, SCHEMA.Place))
                    add((location_bnode, RDFS.label, Literal(location)))

                    # 🔥 Добавляем географические координаты в BNode
                    self._add_location_metadata(location_bnode, sheet_row)
                else:
                    location_bnode = self.fls[location]
                # Связываем образец с локацией
                add((self.sample, SCHEMA.fromLocation, location_bnode))

        # 🔥 НОВОЕ: Обработка координатных данных
        elif field in ["LATITUDE_MIN", "LATITUDE_MAX", "LONGITUDE_MIN", "LONGITUDE_MAX",
                      "LOCATION_COMMENT", "LAND_SEA_SAMPLING", "ELEVATION_MIN", "ELEVATION_MAX"]:
            # Эти поля будут обработаны в _add_location_metadata
            pass

        # 🔥 НОВОЕ: Обработка текстовых свойств
        elif field == "LOCATION_COMMENT":
            # Сохраняем для последующего добавления в BNode локации
            self._current_location_comment = val

        elif field == "LAND_SEA_SAMPLING":
            # Обработка типа отбора проб
            self._process_sampling_type(val)

        elif field == "ALTERATION":
            # Обработка альтерации
            self._process_alteration(val)

        elif field == "ROCK_TEXTURE":
            # Обработка текстуры через нашу систему
            self._process_rock_texture(val)

        elif field == "ROCK_NAME":
            # Обработка названия породы
            self._process_rock_name(val)

        elif field == "TECTONIC_SETTING":
            # Обработка тектонической обстановки
            self._process_tectonic_setting(val)

        elif field == "MINERAL":
            # Обработка минерала
            self._process_mineral(val)

        elif field == "PRIMARY_SECONDARY":
            # Обработка типа включения
            self._process_inclusion_type(val)

        # Обработка химических данных (существующий код)
        elif self.sample is not None:
            self.proc_comp((field + prt, fieldname + prt), cell.value)
        elif detlim:
            self.proc_comp((field + prt, fieldname + prt), cell.value, detlim)
        else:
            value = self.proc_value(cell.value)
            if value is not None and self.kwargs['sheetName'] != 'References':
                print("#! ERROR: nowhere to store {} R:{} C:{}\n#!{}".format(
                    cell, row, col, self.header))
                quit()

    def _add_location_metadata(self, location_bnode, row):
        """Добавляет метаданные локации в BNode"""
        add = self.add

        # Собираем все координатные данные из строки
        location_data = {}

        for col, (field, fieldname) in self.header.items():
            cell = row[col]
            if cell.ctype in [xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK]:
                continue

            val = cell.value

            # 🔥 Координаты и географические данные
            if field == "LATITUDE_MIN":
                location_data['lat_min'] = val
                add((location_bnode, PT.latitudeMin, Literal(val)))
            elif field == "LATITUDE_MAX":
                location_data['lat_max'] = val
                add((location_bnode, PT.latitudeMax, Literal(val)))
            elif field == "LONGITUDE_MIN":
                location_data['long_min'] = val
                add((location_bnode, PT.longitudeMin, Literal(val)))
            elif field == "LONGITUDE_MAX":
                location_data['long_max'] = val
                add((location_bnode, PT.longitudeMax, Literal(val)))
            elif field == "LOCATION_COMMENT":
                add((location_bnode, PT.locationComment, Literal(val)))
            elif field == "ELEVATION_MIN":
                add((location_bnode, PT.elevationMin, Literal(val)))
            elif field == "ELEVATION_MAX":
                add((location_bnode, PT.elevationMax, Literal(val)))

        # 🔥 Создаем точку/полигон на основе координат
        self._create_geometry(location_bnode, location_data)

    def _create_geometry(self, location_bnode, location_data):
        """Создает геометрический объект для локации с правильной онтологией"""
        add = self.add

        # 🔥 Указываем тип для BNode локации
        add((location_bnode, RDF.type, PT.GeoBounds))

        # Если есть и min и max координаты - создаем bounding box
        if all(k in location_data for k in ['lat_min', 'lat_max', 'long_min', 'long_max']):
            bbox = BNode()
            add((location_bnode, SCHEMA.geo, bbox))
            add((bbox, RDF.type, SCHEMA.GeoShape))
            add((bbox, SCHEMA.box, Literal(
                f"{location_data['lat_min']} {location_data['long_min']} "
                f"{location_data['lat_max']} {location_data['long_max']}"
            )))

        # Если только одна пара координат - создаем точку
        elif 'lat_min' in location_data and 'long_min' in location_data:
            point = BNode()
            add((location_bnode, WGS.location, point))
            add((point, RDF.type, WGS.Point))
            add((point, WGS.lat, Literal(location_data['lat_min'])))
            add((point, WGS.long, Literal(location_data['long_min'])))

    def _process_sampling_type(self, value):
        """Обработка типа отбора проб"""
        if not hasattr(self, 'sample') or self.sample is None:
            return

        value = str(value).strip().lower()
        if value in ['subaerial', 'submarine', 'underground']:
            self.add((self.sample, PT.samplingEnvironment, PT[value]))
        else:
            # Создаем литерал для неизвестных типов
            self.add((self.sample, PT.samplingEnvironment, Literal(value)))

    def _process_alteration(self, value):
        """Обработка альтерации породы с созданием сущностей"""
        if not hasattr(self, 'sample') or self.sample is None:
            return

        if isinstance(value, str) and value.strip():
            alterations = [alt.strip().lower() for alt in value.split(",")]
            for alteration in alterations:
                if alteration:
                    alt_iri = PT[normURI(alteration)]

                    # 🔥 Создаем сущность AlterationType если не существует
                    if (alt_iri, RDF.type, PT.AlterationType) not in self.g:
                        self.add((alt_iri, RDF.type, PT.AlterationType))
                        self.add((alt_iri, RDFS.label, Literal(alteration.capitalize())))

                    self.add((self.sample, PT.hasAlteration, alt_iri))

    def _process_rock_texture(self, value):
        """Обработка текстуры породы через нашу систему"""
        if not hasattr(self, 'sample') or self.sample is None:
            return

        if isinstance(value, str) and value.strip():
            texture_iri = get_texture_iri(value)
            if texture_iri:
                self.add((self.sample, PT.rockTexture, texture_iri))
            else:
                # Fallback: создаем литерал
                self.add((self.sample, PT.rockTexture, Literal(value)))

    def _process_rock_name(self, value):
        """Обработка названия породы"""
        if not hasattr(self, 'sample') or self.sample is None:
            return

        if isinstance(value, str) and value.strip():
            rocks = [rock.strip().lower() for rock in value.split(",")]
            for rock in rocks:
                if rock:
                    if rock in ['xenolith']:
                        self.add((self.sample, PT.geologicalStructure, PT[rock]))
                    elif rock in ['megacryst']:
                        self.add((self.sample, PT.geologicUnit, PT[rock]))
                    elif rock in COMMON_MINERALS:  # Предопределенный список минералов
                        self.add((self.sample, PT.mineral, PT[rock]))
                    else:
                        # Предполагаем, что это тип породы
                        rock_iri = PT[normURI(rock)]
                        self.add((rock_iri, RDF.type, PT.RockType))
                        self.add((rock_iri, RDFS.label, Literal(rock.capitalize())))
                        self.add((self.sample, PT.rockType, rock_iri))

    def _process_tectonic_setting(self, value):
        """Обработка тектонической обстановки"""
        if not hasattr(self, 'sample') or self.sample is None:
            return

        if isinstance(value, str) and value.strip():
            orig = capitalize(value)
            val_norm = value.strip().lower()
            val_iri = normURI(val_norm)
            setting_iri = PT[val_iri]
            self.add((setting_iri, RDF.type, PT.TectonicSetting))
            self.add((setting_iri, RDFS.label, Literal(orig)))
            self.add((self.sample, PT.tectonicSetting, setting_iri))

    def _process_mineral(self, value):
        """Обработка минерала"""
        if not hasattr(self, 'sample') or self.analysis is None:
            return

        if isinstance(value, str) and value.strip():
            minerals = [mineral.strip().lower() for mineral in value.split(",")]
            for mineral in minerals:
                if mineral:
                    mineral_iri = PT[normURI(mineral)]
                    self.add((mineral_iri, RDF.type, PT.Mineral))
                    self.add((mineral_iri, RDFS.label, Literal(mineral.capitalize())))
                    self.add((self.analysis, PT.mineral, mineral_iri))

    def _process_inclusion_type(self, value):
        """Обработка типа включения"""
        if not hasattr(self, 'sample') or self.sample is None:
            return

        if isinstance(value, str) and value.strip():
            value = value.strip().lower()
            if "primary" in value:
                inclusion_type = "Primary"
            elif "secondary" in value:
                inclusion_type = "Secondary"
            else:
                inclusion_type = "Primary"  # значение по умолчанию

            self.add((self.sample, PT.inclusion, PT[inclusion_type]))

    def fCITATION(self, value, **kw):
        add = self.add

        add((self.sample, DCTERMS.bibliographicCitation, Literal(value)))

    _field_map_ = {
        'CITATION': fCITATION,
    }

 # 🔥 Предопределенные списки для классификации
COMMON_MINERALS = {
    "garnet", "spinel", "olivine", "clinopyroxene", "orthopyroxene",
    "ilmenite", "phlogopite", "amphibole", "biotite", "chromite",
    "kyanite", "diamond", "graphite", "corundum", "sanidine",
    "enstatite", "fassaite"
}

COMMON_ROCK_TYPES = {
    "lherzolite", "harzburgite", "peridotite", "wehrlite", "dunite",
    "websterite", "clinopyroxenite", "orthopyroxenite", "pyroxenite",
    "eclogite"
}
#

class Alrosa_Xenolites(Alrosa):
    pass


FILES = {
    'БД georock corr_MVG.xls': (Alrosa_Xenolites, {
        'pages': (0, 1)
    }),
    'БД гранаты из ксенолитов.xls': (Alrosa_Xenolites, {
        'pages': []
    })
}


def parse_sheet(sh, sheetIRI, sheetName, comp):
    # print("Cell D30 is {0}".format(sh.cell_value(rowx=29, colx=3)))
    constr, locs = comp
    st = constr(G, sheetIRI, locations=locs, sheetName=sheetName, sheet=sh)
    G.add((sheetIRI, RDF.type, DataSheet))
    sheetName = sheetName.replace(".xls_", ", ")
    G.add((sheetIRI, RDFS.label, Literal(sheetName)))
    print("Parsing sheet: {}".format(sheetName))
    for rx in range(sh.nrows):
        st.row(sh.row(rx), rx)
    #print("PROBLEMATICS:")
    #pprint(st.non_iso)


def parse_xl(file, comp):
    """
    Конвертирует указанный Excel-файл.

    Аргументы:
    - file: Имя файла для обработки
    - comp: Кортеж (класс-обработчик, список_локаций)
    """

    print("# FILE: {} at {}".format(file, SUBDIR))
    pathfile = os.path.join(SUBDIR, file)
    # df = pd.read_excel(pathfile)
    wb = xlrd.open_workbook(pathfile)
    print("# Sheet names: {}".format(wb.sheet_names()))
    for sheet_no, sheet in enumerate(wb.sheet_names()):
        print("# Wb: {}, sheet: {}".format(file, sheet))
        sh = wb.sheet_by_name(sheet)
        sheetname = normURI(file + "_" + sheet)
        print("{0} {1} {2}".format(sh.name, sh.nrows, sh.ncols))
        constr, _ = comp
        whole_name = file + "_" + sheet
        if hasattr(constr, '_sheet_names_'):
            if sheet in constr._sheet_names_:
                parse_sheet(sh, P[sheetname], whole_name, comp)
            elif sheet_no in constr._sheet_names_:
                parse_sheet(sh, P[sheetname], whole_name, comp)
        else:
            parse_sheet(sh, P[sheetname], whole_name, comp)


def update(g):
    # qres=g.query('''
    # PREFIX pt: <http://crust.irk.ru/ontology/pollution/terms/1.0/>
    # PREFIX p: <http://crust.irk.ru/ontology/pollution/1.0/>
    # PREFIX mt: <http://www.daml.org/2003/01/periodictable/PeriodicTable#>
    # PREFIX wgs: <https://www.w3.org/2003/01/geo/wgs84_pos#>
    # SELECT ?sample
    # WHERE {
    #     ?sample a pt:GeologicalSample .
    #     ?sample wgs:long ?long .
    #     ?sample wgs:lat ?lat .
    # }
    # ''')
    # for row in qres:
    #     print(row)

    g.update("""
    PREFIX pt: <http://crust.irk.ru/ontology/pollution/terms/1.0/>
    PREFIX p: <http://crust.irk.ru/ontology/pollution/1.0/>
    PREFIX mt: <http://www.daml.org/2003/01/periodictable/PeriodicTable#>
    PREFIX wgs: <https://www.w3.org/2003/01/geo/wgs84_pos#>
    DELETE {
        ?sample a wgs:SpatialThing .
        ?sample wgs:long ?long .
        ?sample wgs:lat ?lat .
    }
    INSERT {
        ?sample wgs:location _:l .
        _:l a wgs:Point .
        _:l wgs:long ?long .
        _:l wgs:lat ?lat .
    }
    WHERE {
        ?sample a pt:Sample .
        ?sample wgs:long ?long .
        ?sample wgs:lat ?lat .
    }
    """)


PUTURL = "http://ktulhu.isclan.ru:8890/DAV/home/{user}/rdf_sink/{name}"
USER = base64.b64decode("bG9hZGVyCg==").decode("utf8").strip()
CRED = base64.b64decode("bG9hZGVyMzEyCg==").decode("utf8").strip()


def upload(filename, name=None):
    """
    Загружает RDF-файл на сервер.

    Аргументы:
    - filename: Локальное имя файла
    - name: Имя файла на сервере (опционально)
    """
    if name is None:
        name = filename
    with open(filename, "rb") as inp:
        basic = HTTPBasicAuth(USER, CRED)
        URL = PUTURL.format(user=USER, filename=filename, name=name)
        print("URL:{}".format(URL))
        # files = {"file": (filename, inp, "text/turtle")}
        rc = rq.put(
            URL,
            # files=files,
            data=inp.read(),
            auth=basic,
        )
        print("RC:", rc)
        return

        sg = str(P.samples)
        print("URL:", sg)
        data = ({
            "new_name":
            sg,
            "graph_name":
            "http://localhost:8890/DAV/home/loader/rdf_sink/import.ttl",
        }, )
        files = {
            "new_name": (None, sg),
            "graph_name": (
                None,
                "http://localhost:8890/DAV/home/loader/rdf_sink/import.ttl",
            ),
        }
        rc = rq.post(
            "http://ktulhu.isclan.ru:8890/conductor/graphs_page.vspx?page=1",
            files=files,
            # auth=basic
        )
        print("RC:", rc)


if __name__ == "__main__":
    if 1:
        for file, comp in FILES.items():
            parse_xl(file, comp)
            break
        # update(G)
        target = os.path.join(ONTODIR,TARGET)
        with open(target, "w") as o:

            # TODO: Shift location to a BNode using SPARQL.
            # o.write(G.serialize(format='turtle'))
            o.write(G.serialize(format="turtle"))
            print("WROTE: {}".format(target))
    # upload(TARGET, "samples.ttl")
    if 0:
        targetmt = os.path.join(ONTODIR, TARGETMT)
        with open(targetmt, "w") as o:
            o.write(GMT.serialize(format="turtle"))
    print("#!INFO: Normal exit")

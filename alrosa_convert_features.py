import re
from pprint import pprint
from typing import Any, Dict, List, Optional

import pudb
import rdflib
from rdflib import BNode, Graph, Literal, Namespace, URIRef
from rdflib.namespace import RDF, RDFS, XSD

# Определение пространств имен
CRUST = Namespace("http://crust.irk.ru/ontology/contents/terms/1.0/")
RDF_NAMESPACE = RDF
RDFS_NAMESPACE = RDFS
XSD_NAMESPACE = XSD


def normalize_key(key: str) -> str:
    """
    Приведение ключей из сырых данных к каноническому виду.
    Убирает лишние пробелы, переносы строк, приводит к нижнему регистру.
    """
    # Убираем все пробельные символы, включая переносы строк
    normalized = re.sub(r"\s+", " ", key).strip()
    # Заменяем пробелы на подчеркивания для RDF совместимости
    normalized = normalized.replace(" ", "_").replace("-", "_").replace(",", "")
    # Убираем точки в конце
    normalized = normalized.rstrip(".")
    return normalized


def clean_numeric(value: str) -> Optional[float]:
    """
    Преобразует строковое значение в число, обрабатывая запятые и 'н.д.'
    """
    if not value or value == "н.д." or value == "н.д":
        return None

    # Заменяем запятую на точку и убираем пробелы
    cleaned = str(value).replace(",", ".").strip()
    try:
        return float(cleaned)
    except ValueError:
        return None


def extract_type_number_and_mineral(key: str) -> tuple:
    """
    Извлекает номер типа и минерал из ключей вида '1,% gar' или '1, % chr'
    Возвращает (type_number, mineral_code, mineral_uri)
    """
    # Паттерн для поиска номера типа и минерала
    # Ищем цифру в начале, затем необязательные пробелы и запятую, затем код минерала
    match = re.search(r"(\d+),?\s*%\s*(\w+)", key)
    if match:
        type_num = int(match.group(1))
        mineral_code = match.group(2).lower()

        # Маппинг кодов минералов на URI
        mineral_map = {
            "gar": CRUST.garnet,
            "chr": CRUST.chromite,
            "cpx": CRUST.clinopyroxene,
            "ilm": CRUST.ilmenite,
            "ol": CRUST.olivine,
        }

        if mineral_code in mineral_map:
            return type_num, mineral_code, mineral_map[mineral_code]

    return None, None, None


def create_pipe_uri(pipe_id: str) -> URIRef:
    """
    Создает URI для трубки на основе ID
    """
    return CRUST[f"KimberlitePipe/{pipe_id}"]


def convert_features_to_rdf(g: Graph, tube, pipe_uri=None) -> Graph:
    """
    Основной конвертер из словарной структуры в RDF A-Box
    """

    pipe_id, data_dict = tube

    if pipe_uri is None:
        # Создаем URI для трубки
        pipe_uri = create_pipe_uri(pipe_id)
        g.add((pipe_uri, RDF.type, CRUST.KimberlitePipe))

    # Добавляем тип трубки
    g.add((pipe_uri, RDFS.label, Literal(f"Трубка {pipe_id}", lang="ru")))
    g.add((pipe_uri, RDFS.label, Literal(f"Pipe {pipe_id}", lang="en")))
    g.add((pipe_uri, CRUST.pipeId, Literal(pipe_id, datatype=XSD.string)))

    # 1. GEOLOGY - геологические характеристики
    if "geology" in data_dict:
        geo_data = data_dict["geology"]
        geo_bnode = BNode()
        g.add((pipe_uri, CRUST.hasGeology, geo_bnode))
        g.add((geo_bnode, RDF.type, CRUST.PipeGeology))

        # Маппинг ключей геологии на свойства
        geo_mapping = {
            "Форма_тела": CRUST.bodyShape,
            "Возраст_млн_лет": CRUST.ageMillionYears,
            "Перекрытие": CRUST.overburden,
            "Размер": None,  # Обрабатывается отдельно
            "Площадь": CRUST.area,
        }

        for key, prop in geo_mapping.items():
            if key in geo_data and prop:
                value = geo_data[key]
                if key == "Возраст_млн_лет":
                    num_val = clean_numeric(value)
                    if num_val is not None:
                        g.add((geo_bnode, prop, Literal(num_val, datatype=XSD.decimal)))
                elif key == "Площадь":
                    num_val = clean_numeric(value)
                    if num_val is not None:
                        g.add((geo_bnode, prop, Literal(num_val, datatype=XSD.decimal)))
                else:
                    g.add((geo_bnode, prop, Literal(str(value), lang="ru")))

        # Обработка размеров (длина/ширина)
        if "Размер" in geo_data:
            size_str = str(geo_data["Размер"])
            # Предполагаем формат "490 320" или "490/320"
            parts = re.findall(r"\d+", size_str)
            if len(parts) >= 2:
                g.add(
                    (
                        geo_bnode,
                        CRUST.sizeLength,
                        Literal(float(parts[0]), datatype=XSD.decimal),
                    )
                )
                g.add(
                    (
                        geo_bnode,
                        CRUST.sizeWidth,
                        Literal(float(parts[1]), datatype=XSD.decimal),
                    )
                )

    # 2. OLIVINE - петрография оливинов
    if "olivine" in data_dict:
        olivine_data = data_dict["olivine"]
        petro_bnode = BNode()
        g.add((pipe_uri, CRUST.hasPetrography, petro_bnode))
        g.add((petro_bnode, RDF.type, CRUST.Petrography))

        # Маппинг фракций
        fraction_mapping = {
            "1_2_мм": "1_2_mm",
            "2_4_мм": "2_4_mm",
            "4_8_мм": "4_8_mm",
            "8_16_мм": "8_16_mm",
        }

        for fraction_key, fraction_label in fraction_mapping.items():
            if fraction_key in olivine_data:
                value = clean_numeric(olivine_data[fraction_key])
                if value is not None:
                    fraction_bnode = BNode()
                    g.add((petro_bnode, CRUST.hasOlivineFraction, fraction_bnode))
                    g.add((fraction_bnode, RDF.type, CRUST.OlivineSizeFraction))
                    g.add(
                        (
                            fraction_bnode,
                            CRUST.fractionRange,
                            Literal(fraction_label, lang="en"),
                        )
                    )
                    g.add(
                        (
                            fraction_bnode,
                            CRUST.fractionPercentage,
                            Literal(value, datatype=XSD.decimal),
                        )
                    )

    # 3. FEATURES (ABCDE) - целевые показатели
    if "target" in data_dict:
        features_data = data_dict["target"]
        target_bnode = BNode()
        g.add((pipe_uri, CRUST.hasTargetIndicators, target_bnode))
        g.add((target_bnode, RDF.type, CRUST.TargetIndicators))

        feature_mapping = {
            "A": CRUST.paramA,
            "B": CRUST.paramB,
            "C": CRUST.paramC,
            "D": CRUST.paramD,
            "E": CRUST.paramE,
        }

        for key, prop in feature_mapping.items():
            if key in features_data:
                value = features_data[key]
                if prop in [CRUST.paramA, CRUST.paramB, CRUST.paramC, CRUST.paramD]:
                    num_val = clean_numeric(value)
                    if num_val is not None:
                        g.add(
                            (target_bnode, prop, Literal(num_val, datatype=XSD.decimal))
                        )
                else:
                    g.add((target_bnode, prop, Literal(str(value), lang="ru")))

    # 4. ASSOC - алмазная ассоциация и типы минералов
    if "assoc" in data_dict:
        assoc_data = data_dict["assoc"]

        # Создаем основной узел для алмазной ассоциации
        assoc_bnode = BNode()
        g.add((pipe_uri, CRUST.hasDiamondAssociation, assoc_bnode))
        g.add((assoc_bnode, RDF.type, CRUST.DiamondAssociation))

        # 4.1 Классификация гранатов
        garnet_class_bnode = BNode()
        g.add((assoc_bnode, CRUST.hasGarnetClassification, garnet_class_bnode))
        g.add((garnet_class_bnode, RDF.type, CRUST.GarnetDiamondClassification))

        garnet_class_mapping = {
            "алмазная_ассоциация_gar_(по_Соболев_1974)_%_от_перидотитовых_gar_(по_Shulze_2003)": CRUST.garSobolev1974Peridotitic,
            "G10_%": CRUST.g10Percent,
            "G10D_%": CRUST.g10dPercent,
            "G3D_%": CRUST.g3dPercent,
            "G4D_%": CRUST.g4dPercent,
            "G5D_%": CRUST.g5dPercent,
            "Cr2O3_>_5_мас.%_%": CRUST.cr2o3gt5Percent,
            "TiO2_мас.%_(для_перидоти_товых)": CRUST.tio2Peridotitic,
            "TiO2_мас.%_(при_Cr2O3_>_5_мас._%)": CRUST.tio2HighCr,
        }
        pu.db
        for raw_key, prop in garnet_class_mapping.items():
            # Ищем ключ в данных с учетом возможных вариаций
            for data_key, value in assoc_data.items():
                if raw_key in data_key or raw_key.replace(" ", "_") in normalize_key(
                    data_key
                ):
                    num_val = clean_numeric(value)
                    if num_val is not None:
                        g.add(
                            (
                                garnet_class_bnode,
                                prop,
                                Literal(num_val, datatype=XSD.decimal),
                            )
                        )
                    break

        # 4.2 Хромитовая ассоциация
        chromite_assoc_bnode = BNode()
        g.add((assoc_bnode, CRUST.hasChromiteAssociation, chromite_assoc_bnode))
        g.add((chromite_assoc_bnode, RDF.type, CRUST.ChromiteDiamondAssociation))

        chromite_mapping = {
            "Алмазная_ассоциация_%_chr": CRUST.chromiteDiamondPercent,
            "%_принадлежащих_к_перидотитовому_тренду_chr": CRUST.chromitePeridotiticTrendPercent,
        }

        for raw_key, prop in chromite_mapping.items():
            for data_key, value in assoc_data.items():
                if raw_key in data_key or raw_key.replace(" ", "_") in normalize_key(
                    data_key
                ):
                    num_val = clean_numeric(value)
                    if num_val is not None:
                        g.add(
                            (
                                chromite_assoc_bnode,
                                prop,
                                Literal(num_val, datatype=XSD.decimal),
                            )
                        )
                    break

        # 4.3 Ильменитовая классификация
        ilmenite_class_bnode = BNode()
        g.add((assoc_bnode, CRUST.hasIlmeniteClassification, ilmenite_class_bnode))
        g.add((ilmenite_class_bnode, RDF.type, CRUST.IlmeniteClassification))

        ilmenite_mapping = {
            "Кимбер_литовые": CRUST.kimberliticIlmenitePercent,
            "Не_кимбер_литовые": CRUST.nonKimberliticIlmenitePercent,
        }

        for raw_key, prop in ilmenite_mapping.items():
            for data_key, value in assoc_data.items():
                if raw_key in data_key or raw_key.replace(" ", "_") in normalize_key(
                    data_key
                ):
                    num_val = clean_numeric(value)
                    if num_val is not None:
                        g.add(
                            (
                                ilmenite_class_bnode,
                                prop,
                                Literal(num_val, datatype=XSD.decimal),
                            )
                        )
                    break

        # 4.4 Качественные показатели
        quality_bnode = BNode()
        g.add((assoc_bnode, CRUST.hasQualityIndicators, quality_bnode))
        g.add((quality_bnode, RDF.type, CRUST.QualityIndicators))

        quality_mapping = {
            "Хорошая": CRUST.goodQualityPercent,
            "Средняя": CRUST.mediumQualityPercent,
            "Плохая": CRUST.poorQualityPercent,
            "Предельная": CRUST.limitQualityPercent,
            "Отсутствует": CRUST.absentPercent,
        }

        for raw_key, prop in quality_mapping.items():
            for data_key, value in assoc_data.items():
                if raw_key in data_key or raw_key.replace(" ", "_") in normalize_key(
                    data_key
                ):
                    num_val = clean_numeric(value)
                    if num_val is not None:
                        g.add(
                            (
                                quality_bnode,
                                prop,
                                Literal(num_val, datatype=XSD.decimal),
                            )
                        )
                    break

        # 5. МИНЕРАЛЬНЫЕ ТИПЫ - собираем по всем минералам
        mineral_compositions = {}

        for data_key, value in assoc_data.items():
            type_num, mineral_code, mineral_uri = extract_type_number_and_mineral(
                data_key
            )
            if type_num is not None and mineral_uri:
                num_val = clean_numeric(value)
                if num_val is not None:
                    # Создаем композицию для минерала, если еще не создана
                    if mineral_code not in mineral_compositions:
                        comp_bnode = BNode()
                        g.add((pipe_uri, CRUST.mineralComposition, comp_bnode))
                        g.add((comp_bnode, RDF.type, CRUST.MineralTypeComposition))
                        g.add((comp_bnode, CRUST.forMineral, mineral_uri))
                        mineral_compositions[mineral_code] = comp_bnode

                    # Добавляем значение типа
                    value_bnode = BNode()
                    g.add(
                        (
                            mineral_compositions[mineral_code],
                            CRUST.hasTypeValue,
                            value_bnode,
                        )
                    )
                    g.add((value_bnode, RDF.type, CRUST.MineralTypeValue))
                    g.add(
                        (
                            value_bnode,
                            CRUST.typeNumber,
                            Literal(type_num, datatype=XSD.integer),
                        )
                    )
                    g.add(
                        (
                            value_bnode,
                            CRUST.percentage,
                            Literal(num_val, datatype=XSD.decimal),
                        )
                    )

        # 5.1 Добавляем общее количество зерен (n=)
        count_keys = {
            "gar": "(n=_)_gar",
            "chr": "(n=_)_chr",
            "cpx": "(n=_)_Cpx",
            "ilm": "(n=_)_Ilm",
            "ol": "(n=_)_Ol",
        }

        for mineral_code, key_pattern in count_keys.items():
            if mineral_code in mineral_compositions:
                for data_key, value in assoc_data.items():
                    if key_pattern in data_key or key_pattern.replace(
                        " ", "_"
                    ) in normalize_key(data_key):
                        num_val = clean_numeric(value)
                        if num_val is not None:
                            g.add(
                                (
                                    mineral_compositions[mineral_code],
                                    CRUST.totalGrainsCount,
                                    Literal(int(num_val), datatype=XSD.integer),
                                )
                            )
                        break

        # 5.2 Добавляем неизвестные проценты
        unknown_keys = {"chr": "unknown_%_chr", "ilm": "unknown_%_Ilm"}

        for mineral_code, key_pattern in unknown_keys.items():
            if mineral_code in mineral_compositions:
                for data_key, value in assoc_data.items():
                    if key_pattern in data_key or key_pattern.replace(
                        " ", "_"
                    ) in normalize_key(data_key):
                        num_val = clean_numeric(value)
                        if num_val is not None:
                            g.add(
                                (
                                    mineral_compositions[mineral_code],
                                    CRUST.unknownPercent,
                                    Literal(num_val, datatype=XSD.decimal),
                                )
                            )
                        break

        # 5.3 Добавляем LA-ICPMS counts
        la_icpms_keys = {
            "gar": "LA-ICPMS_гранат_N=",
            "cpx": "LA-ICPMS_клинопироксен_N=",
        }

        for mineral_code, key_pattern in la_icpms_keys.items():
            if mineral_code in mineral_compositions:
                for data_key, value in assoc_data.items():
                    if key_pattern in data_key or key_pattern.replace(
                        " ", "_"
                    ) in normalize_key(data_key):
                        num_val = clean_numeric(value)
                        if num_val is not None:
                            g.add(
                                (
                                    mineral_compositions[mineral_code],
                                    CRUST.laIcPmsCount,
                                    Literal(int(num_val), datatype=XSD.integer),
                                )
                            )
                        break

        # 6. ГЕОХИМИЯ МИНЕРАЛОВ
        # 6.1 Геохимия гранатов
        garnet_geo_bnode = BNode()
        g.add((pipe_uri, CRUST.hasMineralGeochemistry, garnet_geo_bnode))
        g.add((garnet_geo_bnode, RDF.type, CRUST.MineralGeochemistry))
        g.add((garnet_geo_bnode, CRUST.forMineralGeo, CRUST.garnet))

        garnet_geo_mapping = {
            "Y_гр/т_(для_перидотитовых)_по_Gar": CRUST.yContentPeridotitic,
            "Y_гр/т_(при_Cr2O3_>_5_мас._%)_по_Gar": CRUST.yContentHighCr,
            "Y_край_Температура_оС_по_Gar": CRUST.yRimTemperature,
        }

        for raw_key, prop in garnet_geo_mapping.items():
            for data_key, value in assoc_data.items():
                if raw_key in data_key or raw_key.replace(" ", "_") in normalize_key(
                    data_key
                ):
                    num_val = clean_numeric(value)
                    if num_val is not None:
                        g.add(
                            (
                                garnet_geo_bnode,
                                prop,
                                Literal(num_val, datatype=XSD.decimal),
                            )
                        )
                    break

        # 7. ГЕОТЕРМАЛЬНЫЕ ДАННЫЕ
        # 7.1 По клинопироксену
        cpx_geo_bnode = BNode()
        g.add((pipe_uri, CRUST.hasGeothermalData, cpx_geo_bnode))
        g.add((cpx_geo_bnode, RDF.type, CRUST.GeothermalData))
        g.add((cpx_geo_bnode, CRUST.geothermalMineral, CRUST.clinopyroxene))

        cpx_geo_mapping = {
            "Геотерма_мВт/м2_по_CPx": CRUST.heatFlow,
            "Мощность_литосферы_км_по_CPx": CRUST.lithosphereThickness,
            "Мощность_области_стабильности_алмаза_км_по_CPx": CRUST.diamondStabilityZone,
        }

        for raw_key, prop in cpx_geo_mapping.items():
            for data_key, value in assoc_data.items():
                if raw_key in data_key or raw_key.replace(" ", "_") in normalize_key(
                    data_key
                ):
                    num_val = clean_numeric(value)
                    if num_val is not None:
                        g.add(
                            (
                                cpx_geo_bnode,
                                prop,
                                Literal(num_val, datatype=XSD.decimal),
                            )
                        )
                    break

        # 7.2 По гранату
        gar_geo_bnode = BNode()
        g.add((pipe_uri, CRUST.hasGeothermalData, gar_geo_bnode))
        g.add((gar_geo_bnode, RDF.type, CRUST.GeothermalData))
        g.add((gar_geo_bnode, CRUST.geothermalMineral, CRUST.garnet))

        gar_geo_mapping = {
            "Геотерма_мВт/м2_по_Gar": CRUST.heatFlow,
            "Мощность_литосферы_км_по_Gar": CRUST.lithosphereThickness,
            "Мощность_алмазного_окна_км_по_Gar": CRUST.diamondWindowThickness,
            "Мощность_области_метасоматоза_км_по_Gar": CRUST.metasomatismZone,
        }

        for raw_key, prop in gar_geo_mapping.items():
            for data_key, value in assoc_data.items():
                if raw_key in data_key or raw_key.replace(" ", "_") in normalize_key(
                    data_key
                ):
                    num_val = clean_numeric(value)
                    if num_val is not None:
                        g.add(
                            (
                                gar_geo_bnode,
                                prop,
                                Literal(num_val, datatype=XSD.decimal),
                            )
                        )
                    break

    return g


# Процедура приведения ключей к каноническому виду
def canonicalize_keys(data_dict: Dict[str, Any]) -> Dict[str, Any]:
    """
    Приводит все ключи в словаре к каноническому виду
    """
    if not isinstance(data_dict, dict):
        return data_dict

    result = {}
    for key, value in data_dict.items():
        if isinstance(value, dict):
            result[normalize_key(key)] = canonicalize_keys(value)
        elif isinstance(value, list):
            result[normalize_key(key)] = [
                canonicalize_keys(item) if isinstance(item, dict) else item
                for item in value
            ]
        else:
            result[normalize_key(key)] = value

    return result

import uuid
from pprint import pprint

from sqlalchemy import (
    JSON,
    Boolean,
    Column,
    DateTime,
    Float,
    ForeignKey,
    Integer,
    String,
    UniqueConstraint,
    create_engine,
)
from sqlalchemy.dialects.postgresql import UUID
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import relationship, sessionmaker
from sqlalchemy.sql import func

Base = declarative_base()


class Diamonds(Base):
    """
    T-Box таблица для данных по алмазам
    """

    __tablename__ = "diamonds"

    # Primary Key
    id = Column(UUID(as_uuid=True), primary_key=True, default=uuid.uuid4)

    # Связь с A-Box (трубка) - одинаковый для всех записей при одноразовом импорте
    pipe_uuid = Column(UUID(as_uuid=True), nullable=False, index=True)

    # Sample identifiers (from source data)
    sample_id = Column(String(50), nullable=False, index=True)  # 'пробы' -> sample_id
    sample_id_alt = Column(String(50), nullable=True)  # 'пробы_1' -> sample_id_alt
    borehole = Column(String(50), nullable=True)  # 'скважина' -> borehole
    rock_type = Column(String(100), nullable=True)  # 'порода' -> rock_type
    interval = Column(String(50), nullable=True)  # 'Интервал' -> interval

    # Weight indicators
    initial_weight_kg = Column(
        Float, nullable=True
    )  # 'Исход_вес_кг' -> initial_weight_kg
    acid_concentrate_kg = Column(
        Float, nullable=True
    )  # 'Выход_кислотного_концентрата_кг' -> acid_concentrate_kg
    salt_concentrate_g = Column(
        Float, nullable=True
    )  # 'Выход_солевого_концентр_г' -> salt_concentrate_g
    alkaline_concentrate_g = Column(
        Float, nullable=True
    )  # 'Выход_щелочного_концентрата_г' -> alkaline_concentrate_g
    heavy_fraction_g = Column(
        Float, nullable=True
    )  # 'Выход_тяжелой_фракции_г' -> heavy_fraction_g
    acid_cleaning_g = Column(
        Float, nullable=True
    )  # 'Кислотная_очистка_солев_Конц_г' -> acid_cleaning_g

    # Diamonds
    diamonds_monocrystals = Column(
        Integer, nullable=True
    )  # 'Количество_обнаруженных_алмазов_монокристаллы' -> diamonds_monocrystals
    diamonds_fragments = Column(
        Integer, nullable=True
    )  # 'Количество_обнаруженных_алмазов_обломки_и_поликр_исталлы' -> diamonds_fragments
    crystals_per_kg = Column(
        Float, nullable=True
    )  # 'кристаллов_обломков_кг' -> crystals_per_kg

    # Service fields (from val_XX substitutions)
    quality_check = Column(
        Boolean, nullable=True
    )  # 'check' (from first val_) -> quality_check
    total_weight = Column(Float, nullable=True)  # 'total' -> total_weight

    # JSONB с фракциями (все диапазоны)
    fractions = Column(JSON, nullable=False, default={})

    # Метаданные
    created_at = Column(DateTime(timezone=True), server_default=func.now())
    updated_at = Column(DateTime(timezone=True), onupdate=func.now())

    def __repr__(self):
        return f"<Diamonds(sample_id='{self.sample_id}', pipe_uuid={self.pipe_uuid})>"

    @classmethod
    def import_from_dataframe(cls, df, pipe_uuid, connection_string, if_exists="fail"):
        """
        Одноразовый импорт данных для конкретной трубки

        Parameters:
        df: pandas DataFrame - предобработанный DataFrame (после preprocess_diamonds)
        pipe_uuid: UUID или str - UUID трубки из A-Box (одинаковый для всех записей)
        connection_string: str - строка подключения к БД
        if_exists: str - что делать если данные уже существуют:
            'fail' - выбросить ошибку
            'replace' - заменить существующие данные
            'append' - добавить к существующим (не рекомендуется для одноразового импорта)

        Returns:
        int - количество загруженных записей
        """

        # Конвертируем pipe_uuid в UUID если передан строкой
        if isinstance(pipe_uuid, str):
            pipe_uuid = uuid.UUID(pipe_uuid)

        # Создаём движок и сессию
        engine = create_engine(connection_string)

        # Создаём таблицу если не существует
        cls.metadata.create_all(engine)

        Session = sessionmaker(bind=engine)
        session = Session()

        try:
            # Проверяем, есть ли уже данные для этой трубки
            existing_count = session.query(cls).filter_by(pipe_uuid=pipe_uuid).count()

            if existing_count > 0:
                if if_exists == "fail":
                    raise ValueError(
                        f"Данные для трубки {pipe_uuid} уже существуют ({existing_count} записей). "
                        f"Используйте if_exists='replace' для замены или 'append' для добавления."
                    )

                elif if_exists == "replace":
                    # Удаляем существующие записи
                    deleted = session.query(cls).filter_by(pipe_uuid=pipe_uuid).delete()
                    session.commit()
                    print(
                        f"Удалено {deleted} существующих записей для трубки {pipe_uuid}"
                    )

                elif if_exists == "append":
                    print(
                        f"Добавление к {existing_count} существующим записям для трубки {pipe_uuid}"
                    )
                else:
                    raise ValueError(f"Недопустимое значение if_exists: {if_exists}")

            # Преобразуем DataFrame в список словарей
            records = df.to_dict("records")
            imported_count = 0

            for record in records:
                # Create model object with field mapping from DataFrame columns
                diamond = cls(
                    pipe_uuid=pipe_uuid,
                    # Map Russian DataFrame columns to English model fields
                    sample_id=record.get(
                        "пробы"
                    ),  # DataFrame: "пробы" -> model: sample_id
                    sample_id_alt=record.get(
                        "пробы_1"
                    ),  # DataFrame: "пробы_1" -> model: sample_id_alt
                    borehole=record.get(
                        "скважина"
                    ),  # DataFrame: "скважина" -> model: borehole
                    rock_type=record.get(
                        "порода"
                    ),  # DataFrame: "порода" -> model: rock_type
                    interval=record.get(
                        "Интервал"
                    ),  # DataFrame: "Интервал" -> model: interval
                    # Weight indicators
                    initial_weight_kg=record.get(
                        "Исход_вес_кг"
                    ),  # DataFrame: "Исход_вес_кг" -> model: initial_weight_kg
                    acid_concentrate_kg=record.get(
                        "Выход_кислотного_концентрата_кг"
                    ),  # DataFrame: "Выход_кислотного_концентрата_кг" -> model: acid_concentrate_kg
                    salt_concentrate_g=record.get(
                        "Выход_солевого_концентр_г"
                    ),  # DataFrame: "Выход_солевого_концентр_г" -> model: salt_concentrate_g
                    alkaline_concentrate_g=record.get(
                        "Выход_щелочного_концентрата_г"
                    ),  # DataFrame: "Выход_щелочного_концентрата_г" -> model: alkaline_concentrate_g
                    heavy_fraction_g=record.get(
                        "Выход_тяжелой_фракции_г"
                    ),  # DataFrame: "Выход_тяжелой_фракции_г" -> model: heavy_fraction_g
                    acid_cleaning_g=record.get(
                        "Кислотная_очистка_солев_Конц_г"
                    ),  # DataFrame: "Кислотная_очистка_солев_Конц_г" -> model: acid_cleaning_g
                    # Diamonds
                    diamonds_monocrystals=record.get(
                        "Количество_обнаруженных_алмазов_монокристаллы"
                    ),  # DataFrame: "Количество_обнаруженных_алмазов_монокристаллы" -> model: diamonds_monocrystals
                    diamonds_fragments=record.get(
                        "Количество_обнаруженных_алмазов_обломки_и_поликр_исталлы"
                    ),  # DataFrame: "Количество_обнаруженных_алмазов_обломки_и_поликр_исталлы" -> model: diamonds_fragments
                    crystals_per_kg=record.get(
                        "кристаллов_обломков_кг"
                    ),  # DataFrame: "кристаллов_обломков_кг" -> model: crystals_per_kg
                    # Service fields (from val_XX substitutions)
                    quality_check=record.get(
                        "check"
                    ),  # DataFrame: "check" (from first val_) -> model: quality_check
                    total_weight=record.get(
                        "total"
                    ),  # DataFrame: "total" -> model: total_weight
                    fractions=record.get(
                        "fractions", {}
                    ),  # DataFrame: "fractions" -> model: fractions
                )
                session.add(diamond)
                imported_count += 1

            session.commit()
            print(
                f"Успешно импортировано {imported_count} записей для трубки {pipe_uuid}"
            )
            return imported_count

        except Exception as e:
            session.rollback()
            print(f"Ошибка при импорте: {e}")
            raise
        finally:
            session.close()

    @classmethod
    def get_by_pipe(cls, pipe_uuid, session):
        """
        Получить все записи для конкретной трубки
        """
        if isinstance(pipe_uuid, str):
            pipe_uuid = uuid.UUID(pipe_uuid)
        return session.query(cls).filter_by(pipe_uuid=pipe_uuid).all()

    @classmethod
    def delete_by_pipe(cls, pipe_uuid, session):
        """
        Удалить все записи для конкретной трубки
        """
        if isinstance(pipe_uuid, str):
            pipe_uuid = uuid.UUID(pipe_uuid)
        deleted = session.query(cls).filter_by(pipe_uuid=pipe_uuid).delete()
        session.commit()
        return deleted


class Sample(Base):
    """
    Сущность Шашка (Sample)
    Связывает шашку с трубкой
    """

    __tablename__ = "samples"

    # Primary Key
    id = Column(UUID(as_uuid=True), primary_key=True, default=uuid.uuid4)

    # Связь с трубкой (A-Box)
    pipe_uuid = Column(UUID(as_uuid=True), nullable=False, index=True)

    # Sample identifier from source data
    sample_name = Column(String(50), nullable=False)  # 'шашка' -> sample_name

    # Sample metadata
    laboratory = Column(String(100), nullable=True)  # 'Лаборатория' -> laboratory
    rock_type = Column(String(100), nullable=True)  # 'Порода' -> rock_type
    depth = Column(String(10), nullable=True)  # 'глубина' -> depth
    class_name = Column(String(50), nullable=True)  # 'класс' -> class_name
    line_borehole = Column(
        String(100), nullable=True
    )  # 'линия_скважина' -> line_borehole
    body = Column(String(100), nullable=True)  # 'тело' -> body
    fraction = Column(String(50), nullable=True)  # 'фракция' -> fraction
    note = Column(String(500), nullable=True)  # 'примечание' -> note
    dimension = Column(String(50), nullable=True)  # 'размерность' -> dimension

    # Связи
    grains = relationship(
        "Grain", back_populates="sample", cascade="all, delete-orphan"
    )

    # Метаданные
    created_at = Column(DateTime(timezone=True), server_default=func.now())
    updated_at = Column(DateTime(timezone=True), onupdate=func.now())

    # Уникальность sample_name в пределах трубки
    __table_args__ = (
        UniqueConstraint("pipe_uuid", "sample_name", name="uix_pipe_sample"),
    )

    def __repr__(self):
        return f"<Sample(sample_name='{self.sample_name}', pipe_uuid={self.pipe_uuid})>"

    @classmethod
    def import_from_dataframe(cls, df, pipe_uuid, connection_string):
        """
        Импорт уникальных шашек из DataFrame EPMA
        """
        engine = create_engine(connection_string)
        cls.metadata.create_all(engine)
        Session = sessionmaker(bind=engine)
        session = Session()

        try:
            # Получаем уникальные шашки
            unique_samples = df[["шашка"]].drop_duplicates()
            imported_count = 0

            for _, row in unique_samples.iterrows():
                sample_name = row["шашка"]

                # Проверяем, существует ли уже
                existing = (
                    session.query(cls)
                    .filter_by(pipe_uuid=pipe_uuid, sample_name=sample_name)
                    .first()
                )

                if not existing:
                    # Собираем метаданные для этой шашки
                    sample_data = df[df["шашка"] == sample_name].iloc[0]

                    sample = cls(
                        pipe_uuid=pipe_uuid,
                        sample_name=sample_name,
                        laboratory=sample_data.get("Лаборатория"),
                        rock_type=sample_data.get("Порода"),
                        depth=sample_data.get("глубина"),
                        class_name=sample_data.get("класс"),
                        line_borehole=sample_data.get("линия_скважина"),
                        body=sample_data.get("тело"),
                        fraction=sample_data.get("фракция"),
                        note=sample_data.get("примечание"),
                        dimension=sample_data.get("размерность"),
                    )
                    session.add(sample)
                    imported_count += 1

            session.commit()
            print(f"Импортировано {imported_count} новых шашек для трубки {pipe_uuid}")

        finally:
            session.close()


class Grain(Base):
    """
    Сущность Зерно (Grain)
    Привязка зерна к шашке
    """

    __tablename__ = "grains"

    id = Column(UUID(as_uuid=True), primary_key=True, default=uuid.uuid4)
    sample_id = Column(
        UUID(as_uuid=True), ForeignKey("samples.id", ondelete="CASCADE"), nullable=False
    )
    grain_name = Column(String(50), nullable=False)

    # Связи
    sample = relationship("Sample", back_populates="grains")
    epma_analyses = relationship(
        "EPMAAnalysis", back_populates="grain", cascade="all, delete-orphan"
    )
    lam_analyses = relationship(
        "LAMAnalysis", back_populates="grain", cascade="all, delete-orphan"
    )

    created_at = Column(DateTime(timezone=True), server_default=func.now())

    __table_args__ = (
        UniqueConstraint("sample_id", "grain_name", name="uix_sample_grain"),
    )


class EPMAAnalysis(Base):
    """
    EPMA анализ зерна
    Содержит все геохимические данные для конкретного зерна
    """

    __tablename__ = "epma_analyses"

    # Primary Key
    id = Column(UUID(as_uuid=True), primary_key=True, default=uuid.uuid4)

    # Связь с зерном
    grain_id = Column(
        UUID(as_uuid=True), ForeignKey("grains.id", ondelete="CASCADE"), nullable=False
    )

    # Основные оксиды
    al2o3 = Column(Float, nullable=True)  # Al2O3
    sio2 = Column(Float, nullable=True)  # SiO2
    tio2 = Column(Float, nullable=True)  # TiO2
    feo = Column(Float, nullable=True)  # FeO
    fe2o3 = Column(Float, nullable=True)  # Fe2O3
    feo_alt = Column(Float, nullable=True)  # FeO_1 (альтернативное измерение)
    mgo = Column(Float, nullable=True)  # MgO
    cao = Column(Float, nullable=True)  # CaO
    na2o = Column(Float, nullable=True)  # Na2O
    k2o = Column(Float, nullable=True)  # K2O
    mno = Column(Float, nullable=True)  # MnO
    p2o5 = Column(Float, nullable=True)  # P2O5
    cr2o3 = Column(Float, nullable=True)  # Cr2O3
    nio = Column(Float, nullable=True)  # NiO
    nio_1 = Column(Float, nullable=True)  # NiO (вариант 2)
    nio_2 = Column(Float, nullable=True)  # NiO (вариант 3)
    coo = Column(Float, nullable=True)  # CoO
    v2o3 = Column(Float, nullable=True)  # V2O3
    v2o3_1 = Column(Float, nullable=True)  # V2O3 (вариант 2)
    zno = Column(Float, nullable=True)  # ZnO
    zno_1 = Column(Float, nullable=True)  # ZnO (вариант 2)

    # Минорные элементы
    v = Column(Float, nullable=True)  # V (металл)
    zn = Column(Float, nullable=True)  # Zn (металл)
    x_coord = Column(Float, nullable=True)  # X координата
    y_coord = Column(Float, nullable=True)  # Y координата

    # Специальные параметры
    t_zn_chr = Column(Float, nullable=True)  # T(Zn) Chr
    total = Column(Float, nullable=True)  # Total
    no = Column(String(20), nullable=True)  # No.

    # Service fields (from val_XX substitutions)
    a_number = Column(Float, nullable=True)  # a_number (from val_20)
    correction = Column(Float, nullable=True)  # correction (from val_17 for tube 1_5)

    # Additional measurement fields (renamed from val_XX)
    measurement_12 = Column(Float, nullable=True)  # val_12 -> measurement_12
    measurement_13 = Column(Float, nullable=True)  # val_13 -> measurement_13
    measurement_14 = Column(Float, nullable=True)  # val_14 -> measurement_14
    measurement_15 = Column(Float, nullable=True)  # val_15 -> measurement_15
    measurement_16 = Column(Float, nullable=True)  # val_16 -> measurement_16
    measurement_17 = Column(
        Float, nullable=True
    )  # val_17 (for tube 2_1) -> measurement_17

    # Counters
    count_akb = Column(Integer, nullable=True)  # 'счет_АКБ' -> count_akb
    count_pk = Column(Integer, nullable=True)  # 'счет_ПК' -> count_pk

    # Additional information
    mineral = Column(String(100), nullable=True)  # 'минерал' -> mineral
    mineral_alt = Column(String(100), nullable=True)  # 'минерал_1' -> mineral_alt
    sum_total = Column(Float, nullable=True)  # 'Сумма' -> sum_total

    # Связь
    grain = relationship("Grain", back_populates="epma_analyses")

    # Метаданные
    created_at = Column(DateTime(timezone=True), server_default=func.now())

    def __repr__(self):
        return f"<EPMAAnalysis(grain_id={self.grain_id})>"

    @classmethod
    def import_from_dataframe(cls, df, pipe_uuid, connection_string):
        """
        Импорт EPMA данных с созданием иерархии Sample -> Grain -> Analysis
        """
        engine = create_engine(connection_string)
        Base.metadata.create_all(engine)
        # pprint(Base.metadata.tables.keys())
        Session = sessionmaker(bind=engine)
        session = Session()

        try:
            # 1. Сначала импортируем шашки
            Sample.import_from_dataframe(df, pipe_uuid, connection_string)

            # 2. Создаем маппинг sample_name -> sample_id
            samples = {s.sample_name: s.id for s in session.query(Sample).all()}

            # 3. Группируем по шашкам и зернам
            grain_cache = {}  # (sample_id, grain_name) -> grain_id
            imported_count = 0

            for _, row in df.iterrows():
                sample_name = row.get("шашка")
                grain_name = row.get("зерно")

                if not sample_name or not grain_name:
                    continue

                sample_id = samples.get(sample_name)
                if not sample_id:
                    continue

                # Получаем или создаем зерно
                grain_key = (sample_id, grain_name)
                if grain_key not in grain_cache:
                    grain = (
                        session.query(Grain)
                        .filter_by(sample_id=sample_id, grain_name=grain_name)
                        .first()
                    )

                    if not grain:
                        grain = Grain(sample_id=sample_id, grain_name=grain_name)
                        session.add(grain)
                        session.flush()  # чтобы получить ID

                    grain_cache[grain_key] = grain.id

                grain_id = grain_cache[grain_key]

                # Создаем анализ
                analysis = cls(
                    grain_id=grain_id,
                    # Основные оксиды
                    al2o3=row.get("Al2O3"),
                    sio2=row.get("SiO2"),
                    tio2=row.get("TiO2"),
                    feo=row.get("FeO"),
                    fe2o3=row.get("Fe2O3"),
                    feo_alt=row.get("FeO_1"),
                    mgo=row.get("MgO"),
                    cao=row.get("CaO"),
                    na2o=row.get("Na2O"),
                    k2o=row.get("K2O"),
                    mno=row.get("MnO"),
                    p2o5=row.get("P2O5"),
                    cr2o3=row.get("Cr2O3"),
                    # Никель (три варианта)
                    nio=row.get("NiO") if "NiO" in row else None,
                    nio_1=row.get("NiO_1") if "NiO_1" in row else None,
                    nio_2=row.get("NiO_2") if "NiO_2" in row else None,
                    coo=row.get("CoO"),
                    # Ванадий (два варианта)
                    v2o3=row.get("V2O3") if "V2O3" in row else None,
                    v2o3_1=row.get("V2O3_1") if "V2O3_1" in row else None,
                    # Цинк (два варианта)
                    zno=row.get("ZnO") if "ZnO" in row else None,
                    zno_1=row.get("ZnO_1") if "ZnO_1" in row else None,
                    # Минорные элементы
                    v=row.get("V"),
                    zn=row.get("Zn"),
                    x_coord=row.get("X"),
                    y_coord=row.get("Y"),
                    # Special parameters
                    t_zn_chr=row.get(
                        "T_Zn_Chr"
                    ),  # DataFrame: "T_Zn_Chr" -> model: t_zn_chr
                    total=row.get("Total"),  # DataFrame: "Total" -> model: total
                    no=row.get("No"),  # DataFrame: "No" -> model: no
                    # Service fields (from val_XX substitutions)
                    a_number=row.get(
                        "a_number"
                    ),  # DataFrame: "a_number" (from val_20) -> model: a_number
                    correction=row.get(
                        "correction"
                    ),  # DataFrame: "correction" (from val_17 for tube 1_5) -> model: correction
                    # Additional measurement fields
                    measurement_12=row.get(
                        "val_12"
                    ),  # DataFrame: "val_12" -> model: measurement_12
                    measurement_13=row.get(
                        "val_13"
                    ),  # DataFrame: "val_13" -> model: measurement_13
                    measurement_14=row.get(
                        "val_14"
                    ),  # DataFrame: "val_14" -> model: measurement_14
                    measurement_15=row.get(
                        "val_15"
                    ),  # DataFrame: "val_15" -> model: measurement_15
                    measurement_16=row.get(
                        "val_16"
                    ),  # DataFrame: "val_16" -> model: measurement_16
                    measurement_17=row.get(
                        "val_17"
                    ),  # DataFrame: "val_17" (for tube 2_1) -> model: measurement_17
                    # Counters
                    count_akb=row.get(
                        "счет_АКБ"
                    ),  # DataFrame: "счет_АКБ" -> model: count_akb
                    count_pk=row.get(
                        "счет_ПК"
                    ),  # DataFrame: "счет_ПК" -> model: count_pk
                    # Minerals
                    mineral=row.get(
                        "минерал"
                    ),  # DataFrame: "минерал" -> model: mineral
                    mineral_alt=row.get(
                        "минерал_1"
                    ),  # DataFrame: "минерал_1" -> model: mineral_alt
                    sum_total=row.get(
                        "Сумма"
                    ),  # DataFrame: "Сумма" -> model: sum_total
                )

                session.add(analysis)
                imported_count += 1

                # Периодический commit для больших DataFrame
                if imported_count % 100 == 0:
                    session.commit()

            session.commit()
            print(
                f"Импортировано {imported_count} EPMA анализов для трубки {pipe_uuid}"
            )

        except Exception as e:
            session.rollback()
            print(f"Ошибка при импорте: {e}")
            raise
        finally:
            session.close()


class LAMAnalysis(Base):
    """
    LAM (LA-ICP-MS) анализ зерна
    Данные лазерной абляции для выбранных зерен
    """

    __tablename__ = "lam_analyses"

    # Primary Key
    id = Column(UUID(as_uuid=True), primary_key=True, default=uuid.uuid4)

    # Связь с зерном (те же зерна, что и в EPMA)
    grain_id = Column(
        UUID(as_uuid=True), ForeignKey("grains.id", ondelete="CASCADE"), nullable=False
    )

    # Основные элементы (породообразующие)
    si = Column(Float, nullable=True)  # Si
    ti = Column(Float, nullable=True)  # Ti
    al = Column(Float, nullable=True)  # Al
    fe = Column(Float, nullable=True)  # Fe
    mn = Column(Float, nullable=True)  # Mn
    mg = Column(Float, nullable=True)  # Mg
    ca = Column(Float, nullable=True)  # Ca
    na = Column(Float, nullable=True)  # Na
    k = Column(Float, nullable=True)  # K
    p = Column(Float, nullable=True)  # P

    # Редкоземельные элементы (REE)
    la = Column(Float, nullable=True)  # La
    ce = Column(Float, nullable=True)  # Ce
    pr = Column(Float, nullable=True)  # Pr
    nd = Column(Float, nullable=True)  # Nd
    sm = Column(Float, nullable=True)  # Sm
    eu = Column(Float, nullable=True)  # Eu
    gd = Column(Float, nullable=True)  # Gd
    tb = Column(Float, nullable=True)  # Tb
    dy = Column(Float, nullable=True)  # Dy
    ho = Column(Float, nullable=True)  # Ho
    er = Column(Float, nullable=True)  # Er
    tm = Column(Float, nullable=True)  # Tm
    yb = Column(Float, nullable=True)  # Yb
    lu = Column(Float, nullable=True)  # Lu

    # Высокозарядные элементы (HFSE)
    zr = Column(Float, nullable=True)  # Zr
    hf = Column(Float, nullable=True)  # Hf
    nb = Column(Float, nullable=True)  # Nb
    ta = Column(Float, nullable=True)  # Ta

    # Крупноионные литофилы (LILE)
    rb = Column(Float, nullable=True)  # Rb
    cs = Column(Float, nullable=True)  # Cs
    ba = Column(Float, nullable=True)  # Ba
    sr = Column(Float, nullable=True)  # Sr

    # Переходные металлы
    sc = Column(Float, nullable=True)  # Sc
    v = Column(Float, nullable=True)  # V
    cr = Column(Float, nullable=True)  # Cr
    co = Column(Float, nullable=True)  # Co
    ni = Column(Float, nullable=True)  # Ni
    cu = Column(Float, nullable=True)  # Cu
    zn = Column(Float, nullable=True)  # Zn

    # Other elements
    ga = Column(Float, nullable=True)  # Ga
    ge = Column(Float, nullable=True)  # Ge (if present)
    arsenic = Column(
        "as", Float, nullable=True
    )  # As (if present) - renamed from 'as' (Python keyword)
    y = Column(Float, nullable=True)  # Y
    sn = Column(Float, nullable=True)  # Sn
    sb = Column(Float, nullable=True)  # Sb (if present)
    pb = Column(Float, nullable=True)  # Pb
    bi = Column(Float, nullable=True)  # Bi (if present)
    th = Column(Float, nullable=True)  # Th
    u = Column(Float, nullable=True)  # U

    # Элементы, которые есть в ваших данных
    be = Column(Float, nullable=True)  # Be
    b = Column(Float, nullable=True)  # B
    li = Column(Float, nullable=True)  # Li

    # Счетчики (если есть в LAM)
    count_akb = Column(Integer, nullable=True)  # 'счет_АКБ'
    count_pk = Column(Integer, nullable=True)  # 'счет_ПК'

    # Дополнительная информация
    rock_type = Column(String(100), nullable=True)  # 'порода' (только для 1_8)

    # Связь с зерном
    grain = relationship("Grain", back_populates="lam_analyses")

    # Метаданные
    created_at = Column(DateTime(timezone=True), server_default=func.now())

    def __repr__(self):
        return f"<LAMAnalysis(grain_id={self.grain_id})>"

    @classmethod
    def import_from_dataframe(cls, df, pipe_uuid, connection_string):
        """
        Импорт LAM данных с привязкой к существующим зернам из EPMA
        """
        engine = create_engine(connection_string)
        Base.metadata.create_all(engine)
        Session = sessionmaker(bind=engine)
        session = Session()

        try:
            # Получаем все зерна для данной трубки с их sample_id
            grains = (
                session.query(Grain.id, Grain.grain_name, Sample.sample_name)
                .join(Sample)
                .filter(Sample.pipe_uuid == pipe_uuid)
                .all()
            )

            # Создаем маппинг (sample_name, grain_name) -> grain_id
            grain_map = {}
            for grain_id, grain_name, sample_name in grains:
                key = (sample_name, grain_name)
                grain_map[key] = grain_id

            imported_count = 0
            skipped_count = 0

            for _, row in df.iterrows():
                sample_name = row.get("шашка")
                grain_name = row.get("зерно")

                if not sample_name or not grain_name:
                    continue

                # Ищем зерно в маппинге
                grain_id = grain_map.get((sample_name, grain_name))

                if not grain_id:
                    # Зерно не найдено в EPMA - пропускаем (экспертный отбор)
                    skipped_count += 1
                    continue

                # Создаем LAM анализ
                analysis = cls(
                    grain_id=grain_id,
                    # Основные элементы
                    si=row.get("Si"),
                    ti=row.get("Ti"),
                    al=row.get("Al"),
                    fe=row.get("Fe"),
                    mn=row.get("Mn"),
                    mg=row.get("Mg"),
                    ca=row.get("Ca"),
                    na=row.get("Na"),
                    k=row.get("K"),
                    p=row.get("P"),
                    # Редкоземельные
                    la=row.get("La"),
                    ce=row.get("Ce"),
                    pr=row.get("Pr"),
                    nd=row.get("Nd"),
                    sm=row.get("Sm"),
                    eu=row.get("Eu"),
                    gd=row.get("Gd"),
                    tb=row.get("Tb"),
                    dy=row.get("Dy"),
                    ho=row.get("Ho"),
                    er=row.get("Er"),
                    tm=row.get("Tm"),
                    yb=row.get("Yb"),
                    lu=row.get("Lu"),
                    # HFSE
                    zr=row.get("Zr"),
                    hf=row.get("Hf"),
                    nb=row.get("Nb"),
                    ta=row.get("Ta"),
                    # LILE
                    rb=row.get("Rb"),
                    cs=row.get("Cs"),
                    ba=row.get("Ba"),
                    sr=row.get("Sr"),
                    # Переходные металлы
                    sc=row.get("Sc"),
                    v=row.get("V"),
                    cr=row.get("Cr"),
                    co=row.get("Co"),
                    ni=row.get("Ni"),
                    cu=row.get("Cu"),
                    zn=row.get("Zn"),
                    # Другие
                    ga=row.get("Ga"),
                    y=row.get("Y"),
                    sn=row.get("Sn"),
                    pb=row.get("Pb"),
                    th=row.get("Th"),
                    u=row.get("U"),
                    be=row.get("Be"),
                    b=row.get("B"),
                    li=row.get("Li"),
                    # Счетчики
                    count_akb=row.get("счет_АКБ"),
                    count_pk=row.get("счет_ПК"),
                    rock_type=row.get("порода"),
                )

                session.add(analysis)
                imported_count += 1

                if imported_count % 100 == 0:
                    session.commit()

            session.commit()
            print(f"Импортировано {imported_count} LAM анализов для трубки {pipe_uuid}")
            print(f"Пропущено {skipped_count} записей (зерна не найдены в EPMA)")

        except Exception as e:
            session.rollback()
            print(f"Ошибка при импорте: {e}")
            raise
        finally:
            session.close()


class Phlogopite(Base):
    """
    Данные по флогопиту (прямая привязка к трубке)
    """

    __tablename__ = "phlogopite"

    # Primary Key
    id = Column(UUID(as_uuid=True), primary_key=True, default=uuid.uuid4)

    # Связь с трубкой (A-Box)
    pipe_uuid = Column(UUID(as_uuid=True), nullable=False, index=True)

    # Идентификаторы из исходных данных
    sample_id = Column(String(50), nullable=True)  # 'Образец'
    point_id = Column(String(50), nullable=True)  # 'точки'
    mineral = Column(String(100), nullable=True)  # 'минерал'
    mineral_alt = Column(String(100), nullable=True)  # 'Минерал' (с заглавной)
    source = Column(String(100), nullable=True)  # 'Источник'
    rock_type = Column(String(100), nullable=True)  # 'Порода'

    # Основные оксиды
    sio2 = Column(Float, nullable=True)  # SiO2
    tio2 = Column(Float, nullable=True)  # TiO2
    al2o3 = Column(Float, nullable=True)  # Al2O3
    feo = Column(Float, nullable=True)  # FeO
    mgo = Column(Float, nullable=True)  # MgO
    cao = Column(Float, nullable=True)  # CaO
    na2o = Column(Float, nullable=True)  # Na2O
    k2o = Column(Float, nullable=True)  # K2O
    mno = Column(Float, nullable=True)  # MnO
    p2o5 = Column(Float, nullable=True)  # P2O5
    cr2o3 = Column(Float, nullable=True)  # Cr2O3
    nio = Column(Float, nullable=True)  # NiO

    # Редкоземельные и другие оксиды
    bao = Column(Float, nullable=True)  # BaO
    sro = Column(Float, nullable=True)  # SrO
    ce2o3 = Column(Float, nullable=True)  # Ce2O3
    la2o3 = Column(Float, nullable=True)  # La2O3
    nd2o3 = Column(Float, nullable=True)  # Nd2O3
    nb2o5 = Column(Float, nullable=True)  # Nb2O5
    ta2o5 = Column(Float, nullable=True)  # Ta2O5
    tho2 = Column(Float, nullable=True)  # ThO2
    so3 = Column(Float, nullable=True)  # SO3

    # Летучие компоненты
    f = Column(Float, nullable=True)  # F
    cl = Column(Float, nullable=True)  # Cl

    # Totals and service fields
    total = Column(Float, nullable=True)  # 'Total' -> total
    measurement_17 = Column(
        Float, nullable=True
    )  # 'val_17' (ignored, left as is) -> measurement_17
    zno = Column(Float, nullable=True)  # 'ZnO' -> zno

    # Метаданные
    created_at = Column(DateTime(timezone=True), server_default=func.now())

    def __repr__(self):
        return f"<Phlogopite(pipe_uuid={self.pipe_uuid}, sample_id='{self.sample_id}')>"

    @classmethod
    def import_from_dataframe(cls, df, pipe_uuid, connection_string, if_exists="fail"):
        """
        Импорт данных флогопита для конкретной трубки
        """
        engine = create_engine(connection_string)
        cls.metadata.create_all(engine)
        Session = sessionmaker(bind=engine)
        session = Session()

        def clean(x):
            if isinstance(x, float):
                return x
            elif isinstance(x, int):
                return x
            return None

        try:
            # Проверяем существующие данные
            existing_count = session.query(cls).filter_by(pipe_uuid=pipe_uuid).count()

            if existing_count > 0:
                if if_exists == "fail":
                    raise ValueError(
                        f"Данные для трубки {pipe_uuid} уже существуют ({existing_count} записей)"
                    )
                elif if_exists == "replace":
                    deleted = session.query(cls).filter_by(pipe_uuid=pipe_uuid).delete()
                    session.commit()
                    print(f"Удалено {deleted} существующих записей")
                elif if_exists == "append":
                    print(f"Добавление к {existing_count} существующим записям")

            # Импортируем новые данные
            records = df.to_dict("records")
            imported_count = 0

            for record in records:
                # pprint(record)
                phlog = cls(
                    pipe_uuid=pipe_uuid,
                    # Идентификаторы
                    sample_id=record.get("Образец"),
                    point_id=record.get("точки"),
                    mineral=record.get("минерал"),
                    mineral_alt=record.get("Минерал"),
                    source=record.get("Источник"),
                    rock_type=record.get("Порода"),
                    # Основные оксиды
                    sio2=clean(record.get("SiO2")),
                    tio2=clean(record.get("TiO2")),
                    al2o3=clean(record.get("Al2O3")),
                    feo=clean(record.get("FeO")),
                    mgo=clean(record.get("MgO")),
                    cao=clean(record.get("CaO")),
                    na2o=clean(record.get("Na2O")),
                    k2o=clean(record.get("K2O")),
                    mno=clean(record.get("MnO")),
                    p2o5=clean(record.get("P2O5")),
                    cr2o3=clean(record.get("Cr2O3")),
                    nio=clean(record.get("NiO")),
                    # Редкоземельные
                    bao=clean(record.get("BaO")),
                    sro=clean(record.get("SrO")),
                    ce2o3=clean(record.get("Ce2O3")),
                    la2o3=clean(record.get("La2O3")),
                    nd2o3=clean(record.get("Nd2O3")),
                    nb2o5=clean(record.get("Nb2O5")),
                    ta2o5=clean(record.get("Ta2O5")),
                    tho2=clean(record.get("ThO2")),
                    so3=clean(record.get("SO3")),
                    # Летучие
                    f=clean(record.get("F")),
                    cl=clean(record.get("Cl")),
                    # Service fields
                    total=clean(
                        record.get("Total")
                    ),  # DataFrame: 'Total' -> model: total
                    measurement_17=clean(
                        record.get("val_17")
                    ),  # DataFrame: 'val_17' -> model: measurement_17
                    zno=clean(record.get("ZnO")),  # DataFrame: 'ZnO' -> model: zno
                )

                session.add(phlog)
                imported_count += 1

                if imported_count % 100 == 0:
                    session.commit()

            session.commit()
            print(
                f"Импортировано {imported_count} записей флогопита для трубки {pipe_uuid}"
            )
            return imported_count

        except Exception as e:
            session.rollback()
            print(f"Ошибка при импорте: {e}")
            raise
        finally:
            session.close()


class Geochemy(Base):
    """
    Геохимия микроэлементов (прямая привязка к трубке)
    """

    __tablename__ = "geochemy"

    # Primary Key
    id = Column(UUID(as_uuid=True), primary_key=True, default=uuid.uuid4)

    # Связь с трубкой (A-Box)
    pipe_uuid = Column(UUID(as_uuid=True), nullable=False, index=True)

    # Идентификаторы из исходных данных
    sample_id = Column(String(50), nullable=True)  # 'Образец'
    sample_interval = Column(String(50), nullable=True)  # 'Образец_интервал_от'
    borehole = Column(String(50), nullable=True)  # 'Скважина'
    rock_type = Column(String(100), nullable=True)  # 'Порода'
    source = Column(String(100), nullable=True)  # 'Источник'
    number = Column(String(50), nullable=True)  # 'п_п' (№ п/п)

    # Литофильные элементы (LILE)
    rb = Column(Float, nullable=True)  # Rb
    cs = Column(Float, nullable=True)  # Cs
    ba = Column(Float, nullable=True)  # Ba
    sr = Column(Float, nullable=True)  # Sr

    # Высокозарядные элементы (HFSE)
    zr = Column(Float, nullable=True)  # Zr
    hf = Column(Float, nullable=True)  # Hf
    nb = Column(Float, nullable=True)  # Nb
    ta = Column(Float, nullable=True)  # Ta
    th = Column(Float, nullable=True)  # Th
    u = Column(Float, nullable=True)  # U

    # Редкоземельные элементы (REE)
    la = Column(Float, nullable=True)  # La
    ce = Column(Float, nullable=True)  # Ce
    pr = Column(Float, nullable=True)  # Pr
    nd = Column(Float, nullable=True)  # Nd
    sm = Column(Float, nullable=True)  # Sm
    eu = Column(Float, nullable=True)  # Eu
    gd = Column(Float, nullable=True)  # Gd
    tb = Column(Float, nullable=True)  # Tb
    dy = Column(Float, nullable=True)  # Dy
    ho = Column(Float, nullable=True)  # Ho
    er = Column(Float, nullable=True)  # Er
    tm = Column(Float, nullable=True)  # Tm
    yb = Column(Float, nullable=True)  # Yb
    lu = Column(Float, nullable=True)  # Lu

    # Переходные металлы
    sc = Column(Float, nullable=True)  # Sc
    v = Column(Float, nullable=True)  # V
    cr = Column(Float, nullable=True)  # Cr
    co = Column(Float, nullable=True)  # Co
    ni = Column(Float, nullable=True)  # Ni
    cu = Column(Float, nullable=True)  # Cu
    zn = Column(Float, nullable=True)  # Zn

    # Другие элементы
    y = Column(Float, nullable=True)  # Y
    ga = Column(Float, nullable=True)  # Ga
    ge = Column(Float, nullable=True)  # Ge (если есть)
    arsenic = Column("as", Float, nullable=True)  # As
    mo = Column(Float, nullable=True)  # Mo
    sn = Column(Float, nullable=True)  # Sn
    sb = Column(Float, nullable=True)  # Sb (если есть)
    pb = Column(Float, nullable=True)  # Pb
    bi = Column(Float, nullable=True)  # Bi (если есть)
    be = Column(Float, nullable=True)  # Be
    li = Column(Float, nullable=True)  # Li

    # Метаданные
    created_at = Column(DateTime(timezone=True), server_default=func.now())

    def __repr__(self):
        return f"<Geochemy(pipe_uuid={self.pipe_uuid}, sample_id='{self.sample_id}')>"

    @classmethod
    def import_from_dataframe(cls, df, pipe_uuid, connection_string, if_exists="fail"):
        """
        Импорт геохимических данных для конкретной трубки
        """
        engine = create_engine(connection_string)
        cls.metadata.create_all(engine)
        Session = sessionmaker(bind=engine)
        session = Session()

        try:
            # Проверяем существующие данные
            existing_count = session.query(cls).filter_by(pipe_uuid=pipe_uuid).count()

            if existing_count > 0:
                if if_exists == "fail":
                    raise ValueError(
                        f"Данные для трубки {pipe_uuid} уже существуют ({existing_count} записей)"
                    )
                elif if_exists == "replace":
                    deleted = session.query(cls).filter_by(pipe_uuid=pipe_uuid).delete()
                    session.commit()
                    print(f"Удалено {deleted} существующих записей")
                elif if_exists == "append":
                    print(f"Добавление к {existing_count} существующим записям")

            # Импортируем новые данные
            records = df.to_dict("records")
            imported_count = 0

            def clean(x):
                if isinstance(x, float):
                    return x
                elif isinstance(x, int):
                    return x
                return None

            for record in records:
                geochem = cls(
                    pipe_uuid=pipe_uuid,
                    # Идентификаторы
                    sample_id=record.get("Образец"),
                    sample_interval=record.get("Образец_интервал_от"),
                    borehole=record.get("Скважина"),
                    rock_type=record.get("Порода"),
                    source=record.get("Источник"),
                    number=record.get("п_п"),
                    # LILE
                    rb=clean(record.get("Rb")),
                    cs=clean(record.get("Cs")),  # если есть
                    ba=clean(record.get("Ba")),
                    sr=clean(record.get("Sr")),
                    # HFSE
                    zr=clean(record.get("Zr")),
                    hf=clean(record.get("Hf")),
                    nb=clean(record.get("Nb")),
                    ta=clean(record.get("Ta")),
                    th=clean(record.get("Th")),
                    u=clean(record.get("U")),
                    # REE
                    la=clean(record.get("La")),
                    ce=clean(record.get("Ce")),
                    pr=clean(record.get("Pr")),
                    nd=clean(record.get("Nd")),
                    sm=clean(record.get("Sm")),
                    eu=clean(record.get("Eu")),
                    gd=clean(record.get("Gd")),
                    tb=clean(record.get("Tb")),
                    dy=clean(record.get("Dy")),
                    ho=clean(record.get("Ho")),
                    er=clean(record.get("Er")),
                    tm=clean(record.get("Tm")),
                    yb=clean(record.get("Yb")),
                    lu=clean(record.get("Lu")),
                    # Переходные
                    sc=clean(record.get("Sc")),
                    v=clean(record.get("V")),
                    cr=clean(record.get("Cr")),
                    co=clean(record.get("Co")),
                    ni=clean(record.get("Ni")),
                    cu=clean(record.get("Cu")),
                    zn=clean(record.get("Zn")),
                    # Другие
                    y=clean(record.get("Y")),
                    ga=clean(record.get("Ga")),
                    arsenic=clean(record.get("As")),
                    mo=clean(record.get("Mo")),
                    sn=clean(record.get("Sn")),
                    pb=clean(record.get("Pb")),
                    be=clean(record.get("Be")),
                    li=clean(record.get("Li")),
                )

                session.add(geochem)
                imported_count += 1

                if imported_count % 100 == 0:
                    session.commit()

            session.commit()
            print(
                f"Импортировано {imported_count} записей геохимии для трубки {pipe_uuid}"
            )
            return imported_count

        except Exception as e:
            session.rollback()
            print(f"Ошибка при импорте: {e}")
            raise
        finally:
            session.close()


class Petrochemy(Base):
    """
    Петрохимические данные (прямая привязка к трубке)
    """

    __tablename__ = "petrochemy"

    # Primary Key
    id = Column(UUID(as_uuid=True), primary_key=True, default=uuid.uuid4)

    # Связь с трубкой (A-Box)
    pipe_uuid = Column(UUID(as_uuid=True), nullable=False, index=True)

    # Идентификаторы из исходных данных
    sample_id = Column(String(50), nullable=True)  # 'Образец'
    sample_interval = Column(String(50), nullable=True)  # 'Образец_интервал_от'
    borehole = Column(String(50), nullable=True)  # 'Скважина'
    rock_type = Column(String(100), nullable=True)  # 'Порода'
    source = Column(String(100), nullable=True)  # 'Источник'
    number = Column(String(50), nullable=True)  # 'п_п' (№ п/п)

    # Основные оксиды (петрогенные элементы)
    sio2 = Column(Float, nullable=True)  # SiO2
    tio2 = Column(Float, nullable=True)  # TiO2
    al2o3 = Column(Float, nullable=True)  # Al2O3
    fe2o3 = Column(Float, nullable=True)  # Fe2O3
    feo_total = Column(Float, nullable=True)  # FeOtotal
    mgo = Column(Float, nullable=True)  # MgO
    cao = Column(Float, nullable=True)  # CaO
    na2o = Column(Float, nullable=True)  # Na2O
    k2o = Column(Float, nullable=True)  # K2O
    mno = Column(Float, nullable=True)  # MnO
    p2o5 = Column(Float, nullable=True)  # P2O5

    # Летучие компоненты
    h2o = Column(Float, nullable=True)  # H2O
    co2 = Column(Float, nullable=True)  # СО2
    f = Column(Float, nullable=True)  # F
    s = Column(Float, nullable=True)  # S
    loi = Column(Float, nullable=True)  # Ппп (потери при прокаливании)

    # Петрохимические индексы и отношения
    fe_num = Column(Float, nullable=True)  # Fe# (Fenum)
    mg_num = Column(Float, nullable=True)  # Mg# (Mgnum)
    k_na = Column(Float, nullable=True)  # K/Na
    na2o_k2o = Column(Float, nullable=True)  # Na2O+K2O
    ic = Column(Float, nullable=True)  # I.C.
    ilm_i = Column(Float, nullable=True)  # Ilm.I

    # Суммы
    total = Column(Float, nullable=True)  # Сумма

    # Service fields
    measurement_21 = Column(
        Float, nullable=True
    )  # 'val_21' (only for tube 2_4) -> measurement_21

    # Метаданные
    created_at = Column(DateTime(timezone=True), server_default=func.now())

    def __repr__(self):
        return f"<Petrochemy(pipe_uuid={self.pipe_uuid}, sample_id='{self.sample_id}')>"

    @classmethod
    def import_from_dataframe(cls, df, pipe_uuid, connection_string, if_exists="fail"):
        """
        Импорт петрохимических данных для конкретной трубки
        """
        engine = create_engine(connection_string)
        cls.metadata.create_all(engine)
        Session = sessionmaker(bind=engine)
        session = Session()

        def clean(x):
            if isinstance(x, float):
                return x
            elif isinstance(x, int):
                return x
            return None

        try:
            # Проверяем существующие данные
            existing_count = session.query(cls).filter_by(pipe_uuid=pipe_uuid).count()

            if existing_count > 0:
                if if_exists == "fail":
                    raise ValueError(
                        f"Данные для трубки {pipe_uuid} уже существуют ({existing_count} записей)"
                    )
                elif if_exists == "replace":
                    deleted = session.query(cls).filter_by(pipe_uuid=pipe_uuid).delete()
                    session.commit()
                    print(f"Удалено {deleted} существующих записей")
                elif if_exists == "append":
                    print(f"Добавление к {existing_count} существующим записям")

            # Импортируем новые данные
            records = df.to_dict("records")
            imported_count = 0

            for record in records:
                petro = cls(
                    pipe_uuid=pipe_uuid,
                    # Идентификаторы
                    sample_id=record.get("Образец"),
                    sample_interval=record.get("Образец_интервал_от"),
                    borehole=record.get("Скважина"),
                    rock_type=record.get("Порода"),
                    source=record.get("Источник"),
                    number=record.get("п_п"),
                    # Основные оксиды
                    sio2=clean(record.get("SiO2")),
                    tio2=clean(record.get("TiO2")),
                    al2o3=clean(record.get("Al2O3")),
                    fe2o3=clean(record.get("Fe2O3")),
                    feo_total=clean(record.get("FeOtotal")),
                    mgo=clean(record.get("MgO")),
                    cao=clean(record.get("CaO")),
                    na2o=clean(record.get("Na2O")),
                    k2o=clean(record.get("K2O")),
                    mno=clean(record.get("MnO")),
                    p2o5=clean(record.get("P2O5")),
                    # Летучие
                    h2o=clean(record.get("H2O")),
                    co2=clean(record.get("СО2")),
                    f=clean(record.get("F")),
                    s=clean(record.get("S")),
                    loi=clean(record.get("Ппп")),
                    # Индексы
                    fe_num=clean(record.get("Fenum")),
                    mg_num=clean(record.get("Mgnum")),
                    k_na=clean(record.get("K_Na")),
                    na2o_k2o=clean(record.get("Na2O_K2O")),
                    ic=clean(record.get("I_C")),
                    ilm_i=clean(record.get("Ilm_I")),
                    # Суммы
                    total=clean(record.get("Сумма")),
                    # Service fields
                    measurement_21=clean(
                        record.get("val_21")
                    ),  # DataFrame: 'val_21' -> model: measurement_21
                )

                session.add(petro)
                imported_count += 1

                if imported_count % 100 == 0:
                    session.commit()

            session.commit()
            print(
                f"Импортировано {imported_count} записей петрохимии для трубки {pipe_uuid}"
            )
            return imported_count

        except Exception as e:
            session.rollback()
            print(f"Ошибка при импорте: {e}")
            raise
        finally:
            session.close()


class Oxides(Base):
    """
    Данные по оксидам (прямая привязка к трубке)
    """

    __tablename__ = "oxides"

    # Primary Key
    id = Column(UUID(as_uuid=True), primary_key=True, default=uuid.uuid4)

    # Связь с трубкой (A-Box)
    pipe_uuid = Column(UUID(as_uuid=True), nullable=False, index=True)

    # Идентификаторы из исходных данных
    sample_id = Column(String(50), nullable=True)  # 'Образец'
    point_id = Column(String(50), nullable=True)  # 'точки'
    mineral = Column(String(100), nullable=True)  # 'минерал'
    source = Column(String(100), nullable=True)  # 'Источник'
    rock_type = Column(String(100), nullable=True)  # 'Порода'

    # Основные оксиды
    sio2 = Column(Float, nullable=True)  # SiO2
    tio2 = Column(Float, nullable=True)  # TiO2
    al2o3 = Column(Float, nullable=True)  # Al2O3
    fe2o3 = Column(Float, nullable=True)  # Fe2O3
    feo = Column(Float, nullable=True)  # FeO
    mgo = Column(Float, nullable=True)  # MgO
    cao = Column(Float, nullable=True)  # CaO
    na2o = Column(Float, nullable=True)  # Na2O
    k2o = Column(Float, nullable=True)  # K2O
    mno = Column(Float, nullable=True)  # MnO
    p2o5 = Column(Float, nullable=True)  # P2O5
    cr2o3 = Column(Float, nullable=True)  # Cr2O3
    nio = Column(Float, nullable=True)  # NiO

    # Редкоземельные и другие оксиды
    bao = Column(Float, nullable=True)  # BaO
    sro = Column(Float, nullable=True)  # SrO
    ce2o3 = Column(Float, nullable=True)  # Ce2O3
    la2o3 = Column(Float, nullable=True)  # La2O3
    nd2o3 = Column(Float, nullable=True)  # Nd2O3
    nb2o5 = Column(Float, nullable=True)  # Nb2O5
    ta2o5 = Column(Float, nullable=True)  # Ta2O5
    tho2 = Column(Float, nullable=True)  # ThO2
    v2o3 = Column(Float, nullable=True)  # V2O3
    zno = Column(Float, nullable=True)  # ZnO
    so3 = Column(Float, nullable=True)  # SO3

    # Летучие компоненты
    f = Column(Float, nullable=True)  # F

    # Totals
    total_oxides = Column(Float, nullable=True)  # 'Total' -> total_oxides

    # Метаданные
    created_at = Column(DateTime(timezone=True), server_default=func.now())

    def __repr__(self):
        return f"<Oxides(pipe_uuid={self.pipe_uuid}, sample_id='{self.sample_id}')>"

    @classmethod
    def import_from_dataframe(cls, df, pipe_uuid, connection_string, if_exists="fail"):
        """
        Импорт данных по оксидам для конкретной трубки
        """
        engine = create_engine(connection_string)
        cls.metadata.create_all(engine)
        Session = sessionmaker(bind=engine)
        session = Session()

        try:
            # Проверяем существующие данные
            existing_count = session.query(cls).filter_by(pipe_uuid=pipe_uuid).count()

            if existing_count > 0:
                if if_exists == "fail":
                    raise ValueError(
                        f"Данные для трубки {pipe_uuid} уже существуют ({existing_count} записей)"
                    )
                elif if_exists == "replace":
                    deleted = session.query(cls).filter_by(pipe_uuid=pipe_uuid).delete()
                    session.commit()
                    print(f"Удалено {deleted} существующих записей")
                elif if_exists == "append":
                    print(f"Добавление к {existing_count} существующим записям")

            # Импортируем новые данные
            records = df.to_dict("records")
            imported_count = 0

            for record in records:
                oxides = cls(
                    pipe_uuid=pipe_uuid,
                    # Идентификаторы
                    sample_id=record.get("Образец"),
                    point_id=record.get("точки"),
                    mineral=record.get("минерал"),
                    source=record.get("Источник"),
                    rock_type=record.get("Порода"),
                    # Основные оксиды
                    sio2=record.get("SiO2"),
                    tio2=record.get("TiO2"),
                    al2o3=record.get("Al2O3"),
                    fe2o3=record.get("Fe2O3"),
                    feo=record.get("FeO"),
                    mgo=record.get("MgO"),
                    cao=record.get("CaO"),
                    na2o=record.get("Na2O"),
                    k2o=record.get("K2O"),
                    mno=record.get("MnO"),
                    p2o5=record.get("P2O5"),
                    cr2o3=record.get("Cr2O3"),
                    nio=record.get("NiO"),
                    # Редкоземельные и другие
                    bao=record.get("BaO"),
                    sro=record.get("SrO"),
                    ce2o3=record.get("Ce2O3"),
                    la2o3=record.get("La2O3"),
                    nd2o3=record.get("Nd2O3"),
                    nb2o5=record.get("Nb2O5"),
                    ta2o5=record.get("Ta2O5"),
                    tho2=record.get("ThO2"),
                    v2o3=record.get("V2O3"),
                    zno=record.get("ZnO"),
                    so3=record.get("SO3"),
                    # Летучие
                    f=record.get("F"),
                    # Totals
                    total_oxides=record.get(
                        "Total", record.get("total", record.get("val_17"))
                    ),  # DataFrame: 'Total' -> model: total_oxides
                )

                session.add(oxides)
                imported_count += 1

                if imported_count % 100 == 0:
                    session.commit()

            session.commit()
            print(
                f"Импортировано {imported_count} записей оксидов для трубки {pipe_uuid}"
            )
            return imported_count

        except Exception as e:
            session.rollback()
            print(f"Ошибка при импорте: {e}")
            raise
        finally:
            session.close()


class Isotopes(Base):
    """
    Изотопные данные (прямая привязка к трубке)
    """

    __tablename__ = "isotopes"

    # Primary Key
    id = Column(UUID(as_uuid=True), primary_key=True, default=uuid.uuid4)

    # Связь с трубкой (A-Box)
    pipe_uuid = Column(UUID(as_uuid=True), nullable=False, index=True)

    # Идентификаторы из исходных данных
    sample_id = Column(String(50), nullable=True)  # 'Образец'
    sample_id_alt = Column(String(50), nullable=True)  # 'Образец_1', 'Образец_2'
    source = Column(String(100), nullable=True)  # 'Источник'

    # Концентрации элементов (ppm)
    rb_ppm = Column(Float, nullable=True)  # Rb (ppm)
    sr_ppm = Column(Float, nullable=True)  # Sr (ppm)
    sm_ppm = Column(Float, nullable=True)  # Sm (ppm)
    nd_ppm = Column(Float, nullable=True)  # Nd (ppm)
    lu_ppm = Column(Float, nullable=True)  # Lu (ppm)
    hf_ppm = Column(Float, nullable=True)  # Hf (ppm)

    # Отношения концентраций
    rb_sr = Column(Float, nullable=True)  # Rb/Sr
    sm_nd = Column(Float, nullable=True)  # Sm/Nd
    lu_hf = Column(Float, nullable=True)  # Lu/Hf (если есть)

    # Обратные величины
    one_nd = Column(Float, nullable=True)  # 1/Nd
    one_sr = Column(Float, nullable=True)  # 1/Sr

    # Изотопные отношения
    nd143_nd144 = Column(Float, nullable=True)  # 143Nd/144Nd
    nd143_nd144_i = Column(Float, nullable=True)  # (143Nd/144Nd)i
    sm147_nd144 = Column(Float, nullable=True)  # 147Sm/144Nd

    sr87_sr86 = Column(Float, nullable=True)  # 87Sr/86Sr
    sr87_sr86_i = Column(Float, nullable=True)  # (87Sr/86Sr)i
    rb87_sr86 = Column(Float, nullable=True)  # 87Rb/86Sr

    hf176_hf177 = Column(Float, nullable=True)  # 176Hf/177Hf
    lu176_hf177 = Column(Float, nullable=True)  # 176Lu/177Hf

    # Эпсилон-нотации
    eps_nd = Column(Float, nullable=True)  # εNd
    eps_sr = Column(Float, nullable=True)  # εSr
    eps_hf = Column(Float, nullable=True)  # εHf

    # Погрешности (2σ)
    sigma_2 = Column(Float, nullable=True)  # 2σ
    sigma_2_1 = Column(Float, nullable=True)  # 2σ_1
    sigma_2_2 = Column(Float, nullable=True)  # 2σ_2
    sigma_2_3 = Column(Float, nullable=True)  # 2σ_3
    sigma_2_4 = Column(Float, nullable=True)  # 2σ_4

    # Возраст
    age_ma = Column(Float, nullable=True)  # 'Возраст_млн'
    age_ma_1 = Column(Float, nullable=True)  # 'Возраст_млн_1'
    age_ma_2 = Column(Float, nullable=True)  # 'Возраст_млн_2'

    # Метаданные
    created_at = Column(DateTime(timezone=True), server_default=func.now())

    def __repr__(self):
        return f"<Isotopes(pipe_uuid={self.pipe_uuid}, sample_id='{self.sample_id}')>"

    @classmethod
    def import_from_dataframe(cls, df, pipe_uuid, connection_string, if_exists="fail"):
        """
        Импорт изотопных данных для конкретной трубки
        """

        def sigma(value):
            if value is None:
                return None
            if isinstance(value, str) and value.startswith("±"):
                return float(value[1:])
            else:
                return float(value)

        engine = create_engine(connection_string)
        cls.metadata.create_all(engine)
        Session = sessionmaker(bind=engine)
        session = Session()

        try:
            # Проверяем существующие данные
            existing_count = session.query(cls).filter_by(pipe_uuid=pipe_uuid).count()

            if existing_count > 0:
                if if_exists == "fail":
                    raise ValueError(
                        f"Данные для трубки {pipe_uuid} уже существуют ({existing_count} записей)"
                    )
                elif if_exists == "replace":
                    deleted = session.query(cls).filter_by(pipe_uuid=pipe_uuid).delete()
                    session.commit()
                    print(f"Удалено {deleted} существующих записей")
                elif if_exists == "append":
                    print(f"Добавление к {existing_count} существующим записям")

            # Импортируем новые данные
            records = df.to_dict("records")
            imported_count = 0

            for record in records:
                isotopes = cls(
                    pipe_uuid=pipe_uuid,
                    # Идентификаторы
                    sample_id=record.get("Образец"),
                    sample_id_alt=record.get("Образец_1") or record.get("Образец_2"),
                    source=record.get("Источник"),
                    # Концентрации
                    rb_ppm=record.get("Rb_ppm"),
                    sr_ppm=record.get("Sr_ppm"),
                    sm_ppm=record.get("Sm_ppm"),
                    nd_ppm=record.get("Nd_ppm"),
                    lu_ppm=record.get("Lu_ppm"),
                    hf_ppm=record.get("Hf_ppm"),
                    # Отношения
                    rb_sr=record.get("Rb_Sr"),
                    sm_nd=record.get("Sm_Nd"),
                    # Обратные
                    one_nd=record.get("1_Nd"),
                    one_sr=record.get("1_Sr"),
                    # Nd изотопы
                    nd143_nd144=record.get("143Nd_144Nd"),
                    nd143_nd144_i=record.get("143Nd_144Nd_i"),
                    sm147_nd144=record.get("147Sm_144Nd"),
                    # Sr изотопы
                    sr87_sr86=record.get("87Sr_86Sr"),
                    sr87_sr86_i=record.get("87Sr_86Sr_i"),
                    rb87_sr86=record.get("87Rb_86Sr"),
                    # Hf изотопы
                    hf176_hf177=record.get("176Hf_177Hf"),
                    lu176_hf177=record.get("176Lu_177Hf"),
                    # Эпсилон
                    eps_nd=record.get("epsNd"),
                    eps_sr=record.get("epsSr"),
                    eps_hf=record.get("epsHf"),
                    # Погрешности
                    sigma_2=sigma(record.get("2σ")),
                    sigma_2_1=sigma(record.get("2σ_1")),
                    sigma_2_2=sigma(record.get("2σ_2")),
                    sigma_2_3=sigma(record.get("2σ_3")),
                    sigma_2_4=sigma(record.get("2σ_4")),
                    # Возраст
                    age_ma=record.get("Возраст_млн"),
                    age_ma_1=record.get("Возраст_млн_1"),
                    age_ma_2=record.get("Возраст_млн_2"),
                )

                session.add(isotopes)
                imported_count += 1

                if imported_count % 100 == 0:
                    session.commit()

            session.commit()
            print(
                f"Импортировано {imported_count} записей изотопии для трубки {pipe_uuid}"
            )
            return imported_count

        except Exception as e:
            session.rollback()
            print(f"Ошибка при импорте: {e}")
            raise
        finally:
            session.close()

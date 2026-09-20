from sqlalchemy import Column, Integer, String, DateTime, Text
from datetime import datetime
from database import Base


class Upload(Base):
    __tablename__ = "uploads"

    id = Column(
        Integer,
        primary_key=True,
        index=True
    )

    upload_uuid = Column(
        String(64),
        index=True,
        nullable=True
    )

    filename = Column(String(255))
    file_format = Column(String(30), default="csv")
    sheet_name = Column(String(100), nullable=True)

    total_rows = Column(Integer, default=0)
    valid_rows = Column(Integer, default=0)
    invalid_rows = Column(Integer, default=0)
    chunk_count = Column(Integer, default=0)

    error_breakdown = Column(Text, nullable=True)
    rule_config = Column(Text, nullable=True)
    storage_path = Column(String(500), nullable=True)

    uploaded_at = Column(
        DateTime,
        default=datetime.utcnow
    )
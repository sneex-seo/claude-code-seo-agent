"""
Вспомогательный модуль для аналитики на основе данных, похожих на выгрузку GSC.

MVP‑реализация без прямой интеграции с Google Search Console:
- Все функции принимают уже подготовленные строки (словари) с полями:
  url, query, clicks, impressions, ctr (опционально), position.
- Логику можно вызывать из FastAPI endpoints для быстрого прототипирования.
"""

from collections import defaultdict, Counter
from typing import Any, Dict, List, Tuple
from urllib.parse import urlparse


GscRow = Dict[str, Any]


def _safe_float(value: Any, default: float = 0.0) -> float:
    """Безопасное приведение к float с fallback."""
    try:
        if value is None:
            return default
        return float(value)
    except (TypeError, ValueError):
        return default


def _safe_non_negative_float(value: Any, default: float = 0.0) -> float:
    """То же, но не позволяет уйти в отрицательные значения."""
    v = _safe_float(value, default=default)
    return v if v >= 0 else 0.0


def _safe_str(value: Any) -> str:
    return "" if value is None else str(value)


def _normalize_query(text: str) -> str:
    """Нормализация поискового запроса для ключей и сравнения."""
    return _safe_str(text).strip().lower()


def _normalize_url_for_key(url: str) -> str:
    """
    Агрессивная нормализация URL для целей агрегации:
    - lower host
    - убираем схему и www
    - отбрасываем query/fragment
    - нормализуем trailing slash
    """
    raw = _safe_str(url).strip()
    if not raw:
        return ""
    parsed = urlparse(raw)
    host = (parsed.netloc or "").lower()
    if host.startswith("www."):
        host = host[4:]
    path = parsed.path or "/"
    if path != "/" and path.endswith("/"):
        path = path[:-1]
    return f"{host}{path}"


def build_dashboard(rows: List[GscRow]) -> Dict[str, Any]:
    """
    Строит простую "панель" по списку строк GSC-подобных данных.

    Ожидаемые ключи в каждой строке:
      - url: str
      - query: str
      - clicks: int/float
      - impressions: int/float
      - position: int/float
      - ctr: (опционально) проценты или доля, будет нормализовано.
    """
    if not rows:
        return {
            "summary": {
                "total_clicks": 0,
                "total_impressions": 0,
                "avg_ctr": 0.0,
                "avg_ctr_unweighted": 0.0,
                "avg_position": 0.0,
                "data_anomalies": {
                    "negative_values": False,
                    "ctr_over_100_percent": False,
                },
            },
            "top_pages": [],
            "top_queries": [],
        }

    total_clicks = 0.0
    total_impressions = 0.0
    weighted_position_sum = 0.0
    ctr_values: List[float] = []

    clicks_by_url: Dict[str, float] = defaultdict(float)
    clicks_by_query: Dict[str, float] = defaultdict(float)

    has_negative = False
    has_ctr_over_100 = False

    for row in rows:
        url = _safe_str(row.get("url"))
        query = _safe_str(row.get("query"))
        raw_clicks = _safe_float(row.get("clicks"))
        raw_impressions = _safe_float(row.get("impressions"))
        raw_position = _safe_float(row.get("position"))
        ctr_raw = row.get("ctr")

        if raw_clicks < 0 or raw_impressions < 0 or raw_position < 0:
            has_negative = True

        clicks = _safe_non_negative_float(raw_clicks)
        impressions = _safe_non_negative_float(raw_impressions)
        position = _safe_non_negative_float(raw_position)

        total_clicks += clicks
        total_impressions += impressions
        if impressions > 0 and position > 0:
            weighted_position_sum += position * impressions

        # Нормализуем CTR: если значение > 1, считаем что это проценты
        if ctr_raw is not None:
            ctr_val = _safe_float(ctr_raw)
            if ctr_val > 1:
                ctr_val = ctr_val / 100.0
            if ctr_val > 1:
                has_ctr_over_100 = True
                ctr_val = 1.0
            if ctr_val >= 0:
                ctr_values.append(ctr_val)
        elif impressions > 0:
            # При отсутствии CTR можно приблизительно посчитать
            approx_ctr = clicks / impressions
            if approx_ctr > 1:
                has_ctr_over_100 = True
                approx_ctr = 1.0
            ctr_values.append(approx_ctr)

        clicks_by_url[url] += clicks
        clicks_by_query[query] += clicks

    # Глобальный CTR по проекту
    avg_ctr = (total_clicks / total_impressions) if total_impressions > 0 else 0.0
    # Невзвешенное среднее по строкам (для отладки)
    avg_ctr_unweighted = sum(ctr_values) / len(ctr_values) if ctr_values else 0.0
    avg_position = (
        (weighted_position_sum / total_impressions) if total_impressions > 0 else 0.0
    )

    top_pages = sorted(
        [{"url": url, "clicks": clicks} for url, clicks in clicks_by_url.items()],
        key=lambda x: x["clicks"],
        reverse=True,
    )[:10]

    top_queries = sorted(
        [{"query": q, "clicks": clicks} for q, clicks in clicks_by_query.items()],
        key=lambda x: x["clicks"],
        reverse=True,
    )[:10]

    return {
        "summary": {
            "total_clicks": round(total_clicks, 2),
            "total_impressions": round(total_impressions, 2),
            "avg_ctr": round(avg_ctr, 4),
            "avg_ctr_unweighted": round(avg_ctr_unweighted, 4),
            "avg_position": round(avg_position, 2),
            "data_anomalies": {
                "negative_values": has_negative,
                "ctr_over_100_percent": has_ctr_over_100,
            },
        },
        "top_pages": top_pages,
        "top_queries": top_queries,
    }


def find_quick_wins(
    rows: List[GscRow],
    min_position: float = 11.0,
    max_position: float = 20.0,
    min_impressions: float = 10.0,
) -> List[Dict[str, Any]]:
    """
    Ищет "quick wins": запросы/страницы в диапазоне позиций [11..20].

    Возвращает список объектов:
      {
        "url": ...,
        "query": ...,
        "position": ...,
        "impressions": ...,
        "clicks": ...,
        "potential_score": ...,
      }
    """
    quick: List[Dict[str, Any]] = []
    for row in rows:
        position = _safe_non_negative_float(row.get("position"))
        impressions = _safe_non_negative_float(row.get("impressions"))
        clicks = _safe_non_negative_float(row.get("clicks"))
        if (
            position < min_position
            or position > max_position
            or impressions < min_impressions
        ):
            continue

        # Простейшая формула потенциала
        # Чем ближе к 11 позиции и чем больше показов, тем выше балл.
        position_factor = max(1.0, max_position + 1.0 - position)
        ctr_current = clicks / impressions if impressions > 0 else 0.0
        ideal_ctr = 0.15  # условный "идеальный" CTR
        ctr_gap = max(0.0, ideal_ctr - ctr_current)
        potential_score = impressions * position_factor * (1.0 + ctr_gap)

        quick.append(
            {
                "url": _safe_str(row.get("url")),
                "query": _safe_str(row.get("query")),
                "position": round(position, 2),
                "impressions": round(impressions, 2),
                "clicks": round(clicks, 2),
                "potential_score": round(potential_score, 2),
            }
        )

    quick.sort(key=lambda x: x["potential_score"], reverse=True)
    return quick


def analyze_queries_cannibalization(rows: List[GscRow]) -> List[Dict[str, Any]]:
    """
    Простая проверка каннибализации:
    для каждого запроса находим, если у него >1 уникального URL.
    """
    urls_by_query: Dict[str, set] = defaultdict(set)
    for row in rows:
        query_norm = _normalize_query(row.get("query"))
        url_norm = _normalize_url_for_key(row.get("url"))
        if query_norm and url_norm:
            urls_by_query[query_norm].add(url_norm)

    result: List[Dict[str, Any]] = []
    for query_norm, urls in urls_by_query.items():
        if len(urls) > 1:
            result.append(
                {
                    "query": query_norm,
                    "urls": sorted(urls),
                    "urls_count": len(urls),
                }
            )

    result.sort(key=lambda x: x["urls_count"], reverse=True)
    return result


def analyze_queries_new_lost(
    previous_rows: List[GscRow],
    current_rows: List[GscRow],
) -> Dict[str, List[Dict[str, Any]]]:
    """
    Анализирует новые и потерянные ключи между двумя периодами.

    previous_rows / current_rows - такие же структуры, как в остальных функциях.
    Для ключей используется нормализованный URL + нормализованный запрос.
    """

    def _key(row: GscRow) -> Tuple[str, str]:
        return _normalize_url_for_key(row.get("url")), _normalize_query(row.get("query"))

    prev_map: Dict[Tuple[str, str], GscRow] = {}
    curr_map: Dict[Tuple[str, str], GscRow] = {}

    for r in previous_rows:
        prev_map[_key(r)] = r
    for r in current_rows:
        curr_map[_key(r)] = r

    new_keywords: List[Dict[str, Any]] = []
    lost_keywords: List[Dict[str, Any]] = []

    for key, row in curr_map.items():
        if key not in prev_map:
            url, query = key
            new_keywords.append(
                {
                    "url": url,
                    "query": query,
                    "position": _safe_float(row.get("position")),
                    "clicks": _safe_float(row.get("clicks")),
                    "impressions": _safe_float(row.get("impressions")),
                }
            )

    for key, row in prev_map.items():
        if key not in curr_map:
            url, query = key
            lost_keywords.append(
                {
                    "url": url,
                    "query": query,
                    "position": _safe_float(row.get("position")),
                    "clicks": _safe_float(row.get("clicks")),
                    "impressions": _safe_float(row.get("impressions")),
                }
            )

    new_keywords.sort(key=lambda x: x["impressions"], reverse=True)
    lost_keywords.sort(key=lambda x: x["impressions"], reverse=True)

    return {
        "new_keywords": new_keywords,
        "lost_keywords": lost_keywords,
    }


def summarize_rows_for_debug(rows: List[GscRow]) -> Dict[str, Any]:
    """
    Небольшая диагностическая сводка: сколько уникальных URL/запросов и т.п.
    Полезно для быстрого тестирования через API.
    """
    urls = set()
    queries = set()
    positions: Counter = Counter()
    for row in rows:
        urls.add(_safe_str(row.get("url")))
        queries.add(_safe_str(row.get("query")))
        position = int(round(_safe_float(row.get("position")))) if row.get("position") is not None else None
        if position is not None and position > 0:
            positions[position] += 1

    return {
        "rows_count": len(rows),
        "unique_urls": len(urls),
        "unique_queries": len(queries),
        "positions_histogram": dict(positions.most_common(20)),
    }


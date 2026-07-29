import argparse
import json
import math
import zlib
from collections import Counter, defaultdict
from pathlib import Path

import numpy as np


BASE_SEED = 20260726
MC_AGGREGATE_ITERATIONS = 100_000
MC_FIELD_ITERATIONS = 50_000
BOOTSTRAP_ITERATIONS = 2_000


def parse_args():
    parser = argparse.ArgumentParser(
        description=(
            "Run aggregate, discipline-level, Monte Carlo, bootstrap, "
            "leave-one-journal-out, and multiple-testing analyses."
        )
    )
    parser.add_argument(
        "--input",
        type=Path,
        default=Path("analysis/results/analysis_records.json"),
        help="Prepared analysis-record JSON file.",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=Path("analysis/results/analysis_results.json"),
        help="Destination JSON file.",
    )
    parser.add_argument("--seed", type=int, default=BASE_SEED)
    parser.add_argument(
        "--aggregate-monte-carlo",
        type=int,
        default=MC_AGGREGATE_ITERATIONS,
    )
    parser.add_argument(
        "--discipline-monte-carlo",
        type=int,
        default=MC_FIELD_ITERATIONS,
    )
    parser.add_argument(
        "--bootstrap-iterations",
        type=int,
        default=BOOTSTRAP_ITERATIONS,
    )
    return parser.parse_args()

POSITION = ["Introduction", "Related work", "Methods", "Results", "Discussion"]
DISTANCE = [
    "Same sentence",
    "Same paragraph",
    "Same section",
    "Same article",
    "Single self-cited article",
]
FUNCTION = [
    "Foundation",
    "Inspiration",
    "Extension",
    "Application",
    "Elaborated Citation",
    "Comparison",
    "Similarity",
    "Affirmation",
    "Related work",
    "Simple mention",
    "Comparison between Related Work",
    "Future work",
    "Further reading",
    "Historical background",
]
DEPTH = ["Deep citation", "Moderate citation", "Shallow citation"]
STRENGTH = ["1", "2", "3", "4", ">=5"]

DIMENSIONS = {
    "Citation position": ("position", POSITION, "mention"),
    "Citation distance": ("distance", DISTANCE, "mention"),
    "Citation function": ("citationFunction", FUNCTION, "mention"),
    "Citation depth": ("citationDepth", DEPTH, "mention"),
    "Citation strength": ("strengthBucket", STRENGTH, "strength"),
}

JOURNAL_NAMES = {
    "G1-1": "Library Hi Tech",
    "G1-2": "Activities, Adaptation & Aging",
    "G1-3": "Engineering Technology & Applied Science Research",
    "G1-4": "Exploratory Animal and Medical Research",
    "G1-5": "Regional Statistics",
    "G2-1": "Information Management",
    "G2-2": "Nature Aging",
    "G2-3": "Engineering",
    "G2-4": "Animal Nutrition",
    "G2-5": "Landscape and Urban Planning",
}

FIELD_NAMES = {
    "1": "Information Science & Library Science",
    "2": "Gerontology",
    "3": "Engineering, Multidisciplinary",
    "4": "Veterinary Sciences",
    "5": "Geography",
}


def deterministic_seed(label):
    return BASE_SEED + zlib.crc32(label.encode("utf-8"))


def log_gamma(z):
    coefficients = [
        676.5203681218851,
        -1259.1392167224028,
        771.3234287776531,
        -176.6150291621406,
        12.507343278686905,
        -0.13857109526572012,
        9.9843695780195716e-6,
        1.5056327351493116e-7,
    ]
    if z < 0.5:
        return math.log(math.pi) - math.log(math.sin(math.pi * z)) - log_gamma(1 - z)
    z -= 1
    x = 0.99999999999980993
    for index, coefficient in enumerate(coefficients):
        x += coefficient / (z + index + 1)
    t = z + len(coefficients) - 0.5
    return 0.5 * math.log(2 * math.pi) + (z + 0.5) * math.log(t) - t + math.log(x)


def gamma_q(a, x):
    if x < 0 or a <= 0:
        return math.nan
    if x == 0:
        return 1.0
    gln = log_gamma(a)
    if x < a + 1:
        ap = a
        total = 1.0 / a
        delta = total
        for _ in range(1, 1001):
            ap += 1
            delta *= x / ap
            total += delta
            if abs(delta) < abs(total) * 1e-14:
                break
        p = total * math.exp(-x + a * math.log(x) - gln)
        return max(0.0, min(1.0, 1.0 - p))
    b = x + 1 - a
    c = 1 / 1e-300
    d = 1 / b
    h = d
    for i in range(1, 1001):
        an = -i * (i - a)
        b += 2
        d = an * d + b
        if abs(d) < 1e-300:
            d = 1e-300
        c = b + an / c
        if abs(c) < 1e-300:
            c = 1e-300
        d = 1 / d
        delta = d * c
        h *= delta
        if abs(delta - 1) < 1e-14:
            break
    return max(0.0, min(1.0, math.exp(-x + a * math.log(x) - gln) * h))


def chi_square_test(counts):
    observed = np.asarray(counts, dtype=float)
    keep = observed.sum(axis=0) > 0
    observed = observed[:, keep]
    row_totals = observed.sum(axis=1)
    col_totals = observed.sum(axis=0)
    n = observed.sum()
    expected = np.outer(row_totals, col_totals) / n
    statistic = float(((observed - expected) ** 2 / expected).sum())
    df = int(observed.shape[1] - 1)
    p = gamma_q(df / 2, statistic / 2)
    return {
        "countsKept": observed.astype(int).tolist(),
        "statistic": statistic,
        "df": df,
        "p": p,
        "n": int(n),
        "cramersV": math.sqrt(statistic / n),
        "expectedLessThan5": int((expected < 5).sum()),
        "expectedCells": int(expected.size),
        "minExpected": float(expected.min()),
    }


def monte_carlo_exact_p(counts, iterations, seed):
    observed = np.asarray(counts, dtype=int)
    keep = observed.sum(axis=0) > 0
    observed = observed[:, keep]
    row_totals = observed.sum(axis=1)
    col_totals = observed.sum(axis=0)
    n = int(observed.sum())
    n1 = int(row_totals[0])
    expected1 = n1 * col_totals / n
    expected2 = row_totals[1] * col_totals / n
    observed_statistic = float(
        (((observed[0] - expected1) ** 2) / expected1).sum()
        + (((observed[1] - expected2) ** 2) / expected2).sum()
    )
    rng = np.random.default_rng(seed)
    extreme = 0
    remaining = iterations
    chunk_size = 10_000
    while remaining > 0:
        size = min(chunk_size, remaining)
        sampled1 = rng.multivariate_hypergeometric(col_totals, n1, size=size)
        sampled2 = col_totals - sampled1
        statistics = (
            ((sampled1 - expected1) ** 2 / expected1).sum(axis=1)
            + ((sampled2 - expected2) ** 2 / expected2).sum(axis=1)
        )
        extreme += int(np.count_nonzero(statistics >= observed_statistic - 1e-12))
        remaining -= size
    return {
        "p": (extreme + 1) / (iterations + 1),
        "iterations": iterations,
        "extreme": extreme,
        "seed": seed,
    }


def mann_whitney(group1, group2):
    x = np.asarray(group1, dtype=float)
    y = np.asarray(group2, dtype=float)
    combined = np.concatenate([x, y])
    order = np.argsort(combined, kind="mergesort")
    ranks = np.empty(len(combined), dtype=float)
    sorted_values = combined[order]
    tie_counts = []
    start = 0
    while start < len(combined):
        end = start + 1
        while end < len(combined) and sorted_values[end] == sorted_values[start]:
            end += 1
        average_rank = (start + 1 + end) / 2
        ranks[order[start:end]] = average_rank
        tie_counts.append(end - start)
        start = end
    rank_sum1 = ranks[: len(x)].sum()
    u1 = float(rank_sum1 - len(x) * (len(x) + 1) / 2)
    product = len(x) * len(y)
    mean_u = product / 2
    n = len(combined)
    tie_term = sum(count**3 - count for count in tie_counts)
    variance = product / 12 * ((n + 1) - tie_term / (n * (n - 1)))
    continuity = 0.5 if u1 != mean_u else 0.0
    z = (abs(u1 - mean_u) - continuity) / math.sqrt(variance)
    p = math.erfc(abs(z) / math.sqrt(2))
    return {
        "u1": u1,
        "u2": float(product - u1),
        "z": -z if u1 < mean_u else z,
        "p": p,
        "n": int(n),
        "n1": int(len(x)),
        "n2": int(len(y)),
        "mean1": float(x.mean()),
        "sd1": float(x.std(ddof=1)),
        "mean2": float(y.mean()),
        "sd2": float(y.std(ddof=1)),
        "rankBiserial": float(1 - 2 * u1 / product),
    }


def count_categories(rows, key, categories):
    counter = Counter(row[key] for row in rows)
    return [counter[category] for category in categories]


def records_for_dimension(all_records, source):
    if source == "strength":
        rows = [row for row in all_records if row["hasStrength"]]
        for row in rows:
            value = row["strength"]
            row["strengthBucket"] = ">=5" if value >= 5 else str(int(value))
        return rows
    return all_records


def categorical_result(rows, key, categories, iterations, seed_label):
    group1 = [row for row in rows if row["group"] == "G1"]
    group2 = [row for row in rows if row["group"] == "G2"]
    counts = [
        count_categories(group1, key, categories),
        count_categories(group2, key, categories),
    ]
    test = chi_square_test(counts)
    test["monteCarlo"] = monte_carlo_exact_p(
        counts, iterations, deterministic_seed(seed_label)
    )
    totals = [sum(counts[0]), sum(counts[1])]
    percentages = [
        [value / totals[0] for value in counts[0]],
        [value / totals[1] for value in counts[1]],
    ]
    return {
        "categories": categories,
        "counts": counts,
        "percentages": percentages,
        "test": test,
    }


def bh_adjust(items):
    ordered = sorted(enumerate(items), key=lambda pair: pair[1]["p"])
    adjusted = [0.0] * len(items)
    running = 1.0
    m = len(items)
    for reverse_index in range(m - 1, -1, -1):
        original_index, item = ordered[reverse_index]
        rank = reverse_index + 1
        running = min(running, item["p"] * m / rank)
        adjusted[original_index] = running
    return adjusted


def cluster_matrix(rows, key, categories, cluster_key, group):
    group_rows = [row for row in rows if row["group"] == group]
    clusters = sorted({row[cluster_key] for row in group_rows})
    category_index = {category: index for index, category in enumerate(categories)}
    matrix = np.zeros((len(clusters), len(categories)), dtype=float)
    cluster_index = {cluster: index for index, cluster in enumerate(clusters)}
    for row in group_rows:
        matrix[cluster_index[row[cluster_key]], category_index[row[key]]] += 1
    return clusters, matrix


def cluster_sum_count(rows, value_key, cluster_key, group):
    group_rows = [row for row in rows if row["group"] == group]
    clusters = sorted({row[cluster_key] for row in group_rows})
    cluster_index = {cluster: index for index, cluster in enumerate(clusters)}
    sums = np.zeros(len(clusters), dtype=float)
    counts = np.zeros(len(clusters), dtype=float)
    for row in group_rows:
        index = cluster_index[row[cluster_key]]
        sums[index] += float(row[value_key])
        counts[index] += 1
    return clusters, sums, counts


def percentile_ci(values):
    lower, upper = np.percentile(values, [2.5, 97.5])
    return [float(lower), float(upper)]


def bootstrap_categorical(rows, key, categories, label):
    output = {}
    for cluster_key, level_name in [("articleId", "article"), ("label", "journal")]:
        _, matrix1 = cluster_matrix(rows, key, categories, cluster_key, "G1")
        _, matrix2 = cluster_matrix(rows, key, categories, cluster_key, "G2")
        rng = np.random.default_rng(deterministic_seed(f"bootstrap-{label}-{level_name}"))
        differences = np.empty((BOOTSTRAP_ITERATIONS, len(categories)), dtype=float)
        proportions1 = matrix1 / matrix1.sum(axis=1, keepdims=True)
        proportions2 = matrix2 / matrix2.sum(axis=1, keepdims=True)
        for iteration in range(BOOTSTRAP_ITERATIONS):
            if level_name == "journal":
                pair_indices = rng.integers(0, len(matrix1), size=len(matrix1))
                differences[iteration] = (
                    proportions1[pair_indices].mean(axis=0)
                    - proportions2[pair_indices].mean(axis=0)
                )
            else:
                sampled1 = matrix1[
                    rng.integers(0, len(matrix1), size=len(matrix1))
                ].sum(axis=0)
                sampled2 = matrix2[
                    rng.integers(0, len(matrix2), size=len(matrix2))
                ].sum(axis=0)
                differences[iteration] = (
                    sampled1 / sampled1.sum() - sampled2 / sampled2.sum()
                )
        output[level_name] = {
            category: percentile_ci(differences[:, index])
            for index, category in enumerate(categories)
        }
    return output


def bootstrap_mean(rows, value_key, label):
    output = {}
    for cluster_key, level_name in [("articleId", "article"), ("label", "journal")]:
        _, sums1, counts1 = cluster_sum_count(rows, value_key, cluster_key, "G1")
        _, sums2, counts2 = cluster_sum_count(rows, value_key, cluster_key, "G2")
        rng = np.random.default_rng(deterministic_seed(f"bootstrap-{label}-{level_name}"))
        differences = np.empty(BOOTSTRAP_ITERATIONS, dtype=float)
        means1 = sums1 / counts1
        means2 = sums2 / counts2
        for iteration in range(BOOTSTRAP_ITERATIONS):
            if level_name == "journal":
                indices1 = rng.integers(0, len(sums1), size=len(sums1))
                indices2 = indices1
                differences[iteration] = (
                    means1[indices1].mean() - means2[indices2].mean()
                )
            else:
                indices1 = rng.integers(0, len(sums1), size=len(sums1))
                indices2 = rng.integers(0, len(sums2), size=len(sums2))
                mean1 = sums1[indices1].sum() / counts1[indices1].sum()
                mean2 = sums2[indices2].sum() / counts2[indices2].sum()
                differences[iteration] = mean1 - mean2
        output[level_name] = percentile_ci(differences)
    return output


args = parse_args()
INPUT = args.input.resolve()
OUTPUT = args.output.resolve()
BASE_SEED = args.seed
MC_AGGREGATE_ITERATIONS = args.aggregate_monte_carlo
MC_FIELD_ITERATIONS = args.discipline_monte_carlo
BOOTSTRAP_ITERATIONS = args.bootstrap_iterations

if not INPUT.is_file():
    raise FileNotFoundError(f"Prepared analysis data not found: {INPUT}")

payload = json.loads(INPUT.read_text(encoding="utf-8"))
records = payload["records"]
strength_records = records_for_dimension(records, "strength")

aggregate = {}
for dimension, (key, categories, source) in DIMENSIONS.items():
    source_rows = strength_records if source == "strength" else records
    aggregate[dimension] = categorical_result(
        source_rows,
        key,
        categories,
        MC_AGGREGATE_ITERATIONS,
        f"aggregate-{dimension}",
    )

similarity1 = [
    row["semanticSimilarity"] for row in records if row["group"] == "G1"
]
similarity2 = [
    row["semanticSimilarity"] for row in records if row["group"] == "G2"
]
aggregate["Semantic similarity"] = mann_whitney(similarity1, similarity2)

discipline = {}
for field, field_name in FIELD_NAMES.items():
    discipline[field] = {"fieldName": field_name, "dimensions": {}}
    field_records = [row for row in records if row["field"] == field]
    field_strength = [row for row in strength_records if row["field"] == field]
    for dimension, (key, categories, source) in DIMENSIONS.items():
        source_rows = field_strength if source == "strength" else field_records
        discipline[field]["dimensions"][dimension] = categorical_result(
            source_rows,
            key,
            categories,
            MC_FIELD_ITERATIONS,
            f"field-{field}-{dimension}",
        )
    field_similarity1 = [
        row["semanticSimilarity"]
        for row in field_records
        if row["group"] == "G1"
    ]
    field_similarity2 = [
        row["semanticSimilarity"]
        for row in field_records
        if row["group"] == "G2"
    ]
    discipline[field]["dimensions"]["Semantic similarity"] = mann_whitney(
        field_similarity1, field_similarity2
    )

leave_one_out = []
for dimension, (key, categories, source) in DIMENSIONS.items():
    source_rows = strength_records if source == "strength" else records
    for excluded_label in JOURNAL_NAMES:
        retained = [row for row in source_rows if row["label"] != excluded_label]
        result = categorical_result(
            retained,
            key,
            categories,
            MC_FIELD_ITERATIONS,
            f"loo-{dimension}-{excluded_label}",
        )
        leave_one_out.append(
            {
                "dimension": dimension,
                "excludedLabel": excluded_label,
                "excludedJournal": JOURNAL_NAMES[excluded_label],
                **result["test"],
            }
        )

bootstrap = []
for dimension, (key, categories, source) in DIMENSIONS.items():
    source_rows = strength_records if source == "strength" else records
    result = aggregate[dimension]
    intervals = bootstrap_categorical(source_rows, key, categories, dimension)
    for index, category in enumerate(categories):
        observed_difference = (
            result["percentages"][0][index] - result["percentages"][1][index]
        )
        article_ci = intervals["article"][category]
        journal_ci = intervals["journal"][category]
        bootstrap.append(
            {
                "dimension": dimension,
                "category": category,
                "differenceG1MinusG2": observed_difference,
                "articleCI": article_ci,
                "articleSignificant": article_ci[0] > 0 or article_ci[1] < 0,
                "journalCI": journal_ci,
                "journalSignificant": journal_ci[0] > 0 or journal_ci[1] < 0,
            }
        )

similarity_bootstrap = bootstrap_mean(
    records, "semanticSimilarity", "Semantic similarity"
)
similarity_difference = (
    aggregate["Semantic similarity"]["mean1"]
    - aggregate["Semantic similarity"]["mean2"]
)
bootstrap.append(
    {
        "dimension": "Semantic similarity",
        "category": "Mean score",
        "differenceG1MinusG2": similarity_difference,
        "articleCI": similarity_bootstrap["article"],
        "articleSignificant": (
            similarity_bootstrap["article"][0] > 0
            or similarity_bootstrap["article"][1] < 0
        ),
        "journalCI": similarity_bootstrap["journal"],
        "journalSignificant": (
            similarity_bootstrap["journal"][0] > 0
            or similarity_bootstrap["journal"][1] < 0
        ),
    }
)

fdr_input = []
for dimension in DIMENSIONS:
    fdr_input.append(
        {
            "dimension": dimension,
            "p": aggregate[dimension]["test"]["p"],
        }
    )
fdr_input.append(
    {
        "dimension": "Semantic similarity",
        "p": aggregate["Semantic similarity"]["p"],
    }
)
adjusted = bh_adjust(fdr_input)
fdr = [
    {
        **item,
        "adjustedP": adjusted[index],
        "significant": adjusted[index] < 0.05,
    }
    for index, item in enumerate(fdr_input)
]
fdr.sort(key=lambda item: item["p"])

binary_strength_counts = [
    [
        aggregate["Citation strength"]["counts"][0][0],
        sum(aggregate["Citation strength"]["counts"][0][1:]),
    ],
    [
        aggregate["Citation strength"]["counts"][1][0],
        sum(aggregate["Citation strength"]["counts"][1][1:]),
    ],
]
binary_strength = chi_square_test(binary_strength_counts)
binary_strength["monteCarlo"] = monte_carlo_exact_p(
    binary_strength_counts,
    MC_AGGREGATE_ITERATIONS,
    deterministic_seed("binary-strength"),
)

per_label = {}
for label in JOURNAL_NAMES:
    mention_rows = [row for row in records if row["label"] == label]
    strength_rows = [row for row in strength_records if row["label"] == label]
    per_label[label] = {
        "journal": JOURNAL_NAMES[label],
        "mentions": len(mention_rows),
        "articles": len({row["articleId"] for row in mention_rows}),
        "strengthOccurrences": len(strength_rows),
    }

strength_anchor_indices = [
    index for index, row in enumerate(records) if row["hasStrength"]
]
strength_anomalies = []
for anchor_position, start in enumerate(strength_anchor_indices):
    end = (
        strength_anchor_indices[anchor_position + 1]
        if anchor_position + 1 < len(strength_anchor_indices)
        else len(records)
    )
    block = records[start:end]
    expected_length = int(records[start]["strength"])
    article_ids = sorted({row["articleId"] for row in block})
    cited_strings = sorted(
        {
            row["selfCitedArticleRaw"]
            for row in block
            if row["selfCitedArticleRaw"] != ""
        }
    )
    flags = []
    if len(block) != expected_length:
        flags.append("block_length_differs_from_strength")
    if len(article_ids) != 1:
        flags.append("block_crosses_article_boundary")
    if len(cited_strings) > 1:
        flags.append("multiple_self_cited_strings_within_block")
    if flags:
        strength_anomalies.append(
            {
                "strengthId": records[start]["strengthId"],
                "sourceRow": records[start]["sourceRow"],
                "label": records[start]["label"],
                "strength": expected_length,
                "blockRows": len(block),
                "articleIds": article_ids,
                "selfCitedStringCount": len(cited_strings),
                "flags": flags,
            }
        )

current_cited = ""
wrong_keys = set()
wrong_per_label = defaultdict(set)
for row in records:
    if row["selfCitedArticleRaw"]:
        current_cited = row["selfCitedArticleRaw"]
    key = (row["label"], row["articleTitle"], current_cited)
    wrong_keys.add(key)
    wrong_per_label[row["label"]].add(key)

position_raw_counts = Counter(row["positionRaw"] for row in records)
qc = {
    "source": payload["source"],
    "sheet": payload["sheet"],
    "range": payload["range"],
    "mentions": len(records),
    "mentionsByGroup": {
        "G1": sum(row["group"] == "G1" for row in records),
        "G2": sum(row["group"] == "G2" for row in records),
    },
    "articles": len({row["articleId"] for row in records}),
    "articlesByGroup": {
        "G1": len({row["articleId"] for row in records if row["group"] == "G1"}),
        "G2": len({row["articleId"] for row in records if row["group"] == "G2"}),
    },
    "strengthOccurrences": len(strength_records),
    "strengthByGroup": {
        "G1": sum(row["group"] == "G1" for row in strength_records),
        "G2": sum(row["group"] == "G2" for row in strength_records),
    },
    "wrongTitlePairDeduplication": {
        "total": len(wrong_keys),
        "byLabel": {
            label: len(wrong_per_label[label]) for label in JOURNAL_NAMES
        },
    },
    "positionRawCounts": dict(sorted(position_raw_counts.items())),
    "functionCategories": sorted(
        {row["citationFunction"] for row in records}
    ),
    "depthCategories": sorted({row["citationDepth"] for row in records}),
    "strengthAnomalyBlocks": len(strength_anomalies),
    "strengthAnomalies": strength_anomalies,
    "perLabel": per_label,
}

results = {
    "metadata": {
        "analysisDate": "2026-07-26",
        "baseSeed": BASE_SEED,
        "monteCarloAggregateIterations": MC_AGGREGATE_ITERATIONS,
        "monteCarloFieldIterations": MC_FIELD_ITERATIONS,
        "bootstrapIterations": BOOTSTRAP_ITERATIONS,
        "groupDifferenceDirection": "G1 minus G2",
    },
    "qc": qc,
    "aggregate": aggregate,
    "discipline": discipline,
    "binaryStrengthSensitivity": binary_strength,
    "bootstrap": bootstrap,
    "leaveOneJournalOut": leave_one_out,
    "fdr": fdr,
}

OUTPUT.parent.mkdir(parents=True, exist_ok=True)
OUTPUT.write_text(json.dumps(results, ensure_ascii=False, indent=2), encoding="utf-8")

print(
    json.dumps(
        {
            "output": str(OUTPUT),
            "mentions": qc["mentions"],
            "articles": qc["articles"],
            "strengthOccurrences": qc["strengthOccurrences"],
            "wrongDeduplication": qc["wrongTitlePairDeduplication"]["total"],
            "aggregateStrength": aggregate["Citation strength"]["test"],
            "binaryStrength": binary_strength,
            "strengthAnomalyBlocks": qc["strengthAnomalyBlocks"],
        },
        ensure_ascii=False,
        indent=2,
    )
)

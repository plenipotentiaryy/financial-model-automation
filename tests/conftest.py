from copy import deepcopy

import pytest

from finmodel.config import ALIASES, METRICS, TEMPLATE_MAP
from finmodel.samples import ACME, initialize, write_csv
from finmodel.service import load_dataset


@pytest.fixture
def project(tmp_path):
    initialize(tmp_path)
    return tmp_path


@pytest.fixture
def config():
    return deepcopy({"aliases": ALIASES, "template_map": TEMPLATE_MAP, "metric_definitions": METRICS})


@pytest.fixture
def make_company(tmp_path, config):
    def make(rows=None):
        path = tmp_path / "input.csv"
        write_csv(path, deepcopy(ACME) if rows is None else rows)
        dataset = load_dataset([path], config)
        return dataset.companies["acme"]
    return make

"""Bar-Tender charts: genuine Matplotlib SVG figures, rendered locally via Pyodide."""
import io
import base64
import json
import math
import traceback
from types import MappingProxyType
import matplotlib
matplotlib.use('agg')
import matplotlib.pyplot as plt
import numpy as np


def apply_appearance(code, fig, ax, data, categories, series, title):
    if not code.strip():
        return
    # Scripts only see a detached, immutable snapshot of the data. The editor's
    # source state remains in JavaScript; no data writes are returned here.
    records = tuple(MappingProxyType({**point, 'raw': tuple(point['raw']) if point.get('raw') is not None else None}) for point in data)
    namespace = {'fig': fig, 'ax': ax, 'bars': tuple(ax.patches), 'np': np, 'plt': plt,
                 'data': records, 'categories': tuple(categories), 'series': tuple(series)}
    try:
        exec(compile(code, 'appearance.py', 'exec'), namespace, namespace)
    except Exception as error:
        frames = [frame for frame in traceback.extract_tb(error.__traceback__) if frame.filename == 'appearance.py']
        line = getattr(error, 'lineno', None) or (frames[-1].lineno if frames else None)
        location = f' at line {line}' if line is not None else ''
        raise ValueError(f'Appearance script for "{title}"{location}: {type(error).__name__}: {error}') from None


def render_charts(payload_json):
    payload = json.loads(payload_json)
    graphs = payload['graphs']
    combined = payload['combined']
    columns = min(3, len(graphs)) if combined else 1
    rows = math.ceil(len(graphs) / columns)
    width = 5.0 * columns if combined else (5.5 if graphs[0]['width'] == 'narrow' else 9.5)
    height = 4.0 * rows if combined else 4.5
    with plt.rc_context({'text.usetex': False, 'text.parse_math': False, 'svg.fonttype': 'path',
                         'axes.grid': False, 'figure.facecolor': 'white', 'axes.facecolor': 'white'}):
        fig, axes = plt.subplots(rows, columns, figsize=(width, height), squeeze=False,
                                 layout='constrained')
        try:
            summaries = []
            for ax, graph in zip(axes.flat, graphs):
                data = graph['data']
                show_points = graph.get('points', False)
                rng = np.random.default_rng(0)
                # Recompute raw statistics in Python; supplied mean/SD remain as entered.
                for point in data:
                    if point.get('raw') is not None:
                        samples = np.asarray(point['raw'], dtype=float)
                        point['mean'] = float(np.mean(samples))
                        point['sd'] = float(np.std(samples, ddof=1)) if len(samples) > 1 else None
                categories = list(dict.fromkeys(point['category'] for point in data))
                series = list(dict.fromkeys(point['series'] for point in data))
                if graph['average']:
                    if any(point.get('raw') is None for point in data):
                        raise ValueError('Overall average requires raw data in every row.')
                    if 'Overall average' in categories:
                        raise ValueError('Rename the category “Overall average” to use average bars.')
                    for group in series:
                        samples = np.concatenate([point['raw'] for point in data if point['series'] == group])
                        data.append({'category': 'Overall average', 'series': group,
                                     'mean': float(np.mean(samples)), 'raw': samples.tolist(),
                                     'sd': float(np.std(samples, ddof=1)) if len(samples) > 1 else None})
                    categories.append('Overall average')
                positions = {category: i for i, category in enumerate(categories)}
                bar_width = .8 / len(series)
                palette = graph['paletteColors']
                low, high = 0., 0.
                for j, group in enumerate(series):
                    points = [point for point in data if point['series'] == group]
                    x = [positions[point['category']] + (j - (len(series)-1)/2)*bar_width for point in points]
                    values = [point['mean'] for point in points]
                    bars = ax.bar(x, values, bar_width, label=group if graph['grouped'] else None,
                                  color=palette[j % len(palette)])
                    for cx, point, bar in zip(x, points, bars):
                        deviation = point['sd'] if graph['sd'] and point['sd'] is not None else 0
                        low = min(low, point['mean'] - deviation)
                        high = max(high, point['mean'] + deviation)
                        samples = point.get('raw') if show_points else None
                        if samples is not None and len(samples):
                            low, high = min(low, min(samples)), max(high, max(samples))
                            offsets = rng.permutation(np.linspace(-bar_width*.28, bar_width*.28, len(samples))) if len(samples) > 1 else np.array([0.])
                            ax.scatter(cx+offsets, samples, s=20 if not combined else 14,
                                       facecolors='white', edgecolors='black', linewidths=.7,
                                       alpha=.9, zorder=4)
                            bar.set_alpha(.5)
                        if point['category'] == 'Overall average':
                            bar.set_alpha(.65)
                        if graph['sd'] and point['sd'] is not None:
                            ax.errorbar(cx, point['mean'], yerr=point['sd'], fmt='none',
                                        ecolor='black', capsize=3, elinewidth=1.2)
                        if graph['labels']:
                            negative = point['mean'] < 0
                            label_edge = point['mean'] - deviation if negative else point['mean'] + deviation
                            if samples:
                                label_edge = min(label_edge, min(samples)) if negative else max(label_edge, max(samples))
                            ax.annotate(f"{point['mean']:.2f}".rstrip('0').rstrip('.'),
                                        (cx, label_edge),
                                        xytext=(0, -4 if negative else 4), textcoords='offset points',
                                        ha='center', va='top' if negative else 'bottom', fontsize=8 if combined else 9)
                for key in ['lower', 'upper']:
                    if graph[key] != '':
                        limit = float(graph[key])
                        low, high = min(low, limit), max(high, limit)
                        ax.axhline(limit, color='#D62728', linestyle='--', linewidth=1.2)
                span = high-low or 1
                ymin = float(graph['ymin']) if graph['ymin'] != '' else (low-.05*span if low < 0 else 0)
                ymax = float(graph['ymax']) if graph['ymax'] != '' else high + (.15 if graph['labels'] else .05)*span
                if ymax <= ymin:
                    raise ValueError('Y maximum must be greater than Y minimum.')
                ax.set_ylim(ymin, ymax)
                ax.set_xticks(range(len(categories)), categories)
                ax.set_xlabel(graph['xlabel'])
                ax.set_ylabel(graph['ylabel'])
                ax.set_title(graph['title'])
                if graph['grouped']:
                    ax.legend(loc='best', fontsize=8 if combined else 9)
                if combined:
                    ax.tick_params(labelsize=8)
                apply_appearance(graph.get('appearanceCode', ''), fig, ax, data, categories, series, graph['title'])
                summaries.append({'bars': len(data), 'cats': len(categories), 'groups': len(series)})
            for ax in list(axes.flat)[len(graphs):]:
                ax.set_visible(False)
            output = io.StringIO()
            fig.savefig(output, format='svg', metadata={'Creator': f'Bar-Tender · Matplotlib {matplotlib.__version__}'})
            png = io.BytesIO()
            fig.savefig(png, format='png', dpi=200, facecolor='white', transparent=False,
                        metadata={'Software': f'Bar-Tender · Matplotlib {matplotlib.__version__}'})
            return json.dumps({'svg': output.getvalue(), 'png': base64.b64encode(png.getvalue()).decode('ascii'), 'summaries': summaries,
                               'matplotlibVersion': matplotlib.__version__})
        finally:
            plt.close(fig)
            plt.close('all')

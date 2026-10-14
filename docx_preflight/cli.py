"""CLI deliberately avoids reflecting user paths and document values."""

import argparse
import json
import sys

from .scanner import inspect_docx


class PrivateParser(argparse.ArgumentParser):
    def error(self, message):
        # argparse's usual error includes rejected arguments, possibly private paths.
        self.print_usage(sys.stderr)
        self.exit(2, 'Invalid command arguments. Use --help.\n')


def main(argv=None):
    parser = PrivateParser(description='Read-only DOCX checks. No safety guarantee.')
    parser.add_argument('document', help='Local DOCX input; never uploaded or modified.')
    parser.add_argument('--json', action='store_true', help='Machine-readable fixed codes and counts.')
    parser.add_argument('--fail-on', choices=('warning', 'info', 'none'), default='warning',
                        help='Finding threshold for exit 1 (default: warning). Errors always exit 2.')
    args = parser.parse_args(argv)
    report = inspect_docx(args.document)
    if args.json:
        print(json.dumps(report, ensure_ascii=True, sort_keys=True))
    else:
        print('DOCX preflight: ' + report['status'])
        print(report['scope'])
        for finding in report['findings']:
            print('{severity}: {code} ({count})'.format(**finding))
    if report['status'] == 'error':
        return 2
    if args.fail_on == 'info' and report['findings']:
        return 1
    if args.fail_on == 'warning' and report['status'] == 'warnings':
        return 1
    return 0

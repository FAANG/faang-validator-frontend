"""
Shared validation helpers for extracting errors and warnings from validation records.
Used by Samples, Experiments, and Analysis tabs.
"""
import re
from typing import Dict, List, Tuple, Any


def get_all_errors_and_warnings(record: Dict[str, Any]) -> Tuple[Dict[str, List], Dict[str, Any]]:
    """
    Extract all errors and warnings from a validation record.

    Returns:
        Tuple of (errors dict, warnings dict). Each maps field name to list of messages.
    """
    errors: Dict[str, List] = {}
    warnings: Dict[str, Any] = {}

    # From 'errors' object
    if 'errors' in record and record['errors']:
        # Handle errors.errors array (e.g., "Geographic Location: Field required")
        if 'errors' in record['errors'] and isinstance(record['errors']['errors'], list):
            for error_msg in record['errors']['errors']:
                if ':' in error_msg:
                    parts = error_msg.split(':', 1)
                    field = parts[0].strip()
                    message = parts[1].strip() if len(parts) > 1 else error_msg
                    if field not in errors:
                        errors[field] = []
                    errors[field].append(message)
                else:
                    if 'general' not in errors:
                        errors['general'] = []
                    errors['general'].append(error_msg)

        if 'field_errors' in record['errors']:
            for field, messages in record['errors']['field_errors'].items():
                if field not in errors:
                    errors[field] = []
                if isinstance(messages, list):
                    errors[field].extend(messages)
                else:
                    errors[field].append(messages)
        if 'relationship_errors' in record['errors']:
            for message in record['errors']['relationship_errors']:
                field_to_blame = 'general'
                data = record.get('data', {})
                if 'Child Of' in data and data.get('Child Of'):
                    field_to_blame = 'Child Of'
                elif 'Derived From' in data and data.get('Derived From'):
                    field_to_blame = 'Derived From'
                if field_to_blame not in warnings:
                    warnings[field_to_blame] = []
                warnings[field_to_blame].append(message)

    # From 'field_warnings'
    if 'field_warnings' in record and record['field_warnings']:
        for field, messages in record['field_warnings'].items():
            warnings[field] = messages

    # From 'ontology_warnings'
    if 'ontology_warnings' in record and record['ontology_warnings']:
        for message in record['ontology_warnings']:
            match = re.search(r"in field '([^']*)'", message)
            if match:
                field = match.group(1)
                if field not in warnings:
                    warnings[field] = []
                warnings[field].append(message)
            else:
                if 'general' not in warnings:
                    warnings['general'] = []
                warnings['general'].append(message)

    # From 'relationship_errors' (top level - treat as warnings)
    if 'relationship_errors' in record and record['relationship_errors']:
        field_to_blame = 'general'
        data = record.get('data', {})
        if 'Child Of' in data and data.get('Child Of'):
            field_to_blame = 'Child Of'
        elif 'Derived From' in data and data.get('Derived From'):
            field_to_blame = 'Derived From'
        if field_to_blame not in warnings:
            warnings[field_to_blame] = []
        warnings[field_to_blame].extend(record['relationship_errors'])

    return errors, warnings

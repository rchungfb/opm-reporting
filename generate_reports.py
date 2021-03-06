"""Simple script to build weekly reporting for ISOS team.

Script will pull data from quip.com and format locally in various formats.
This is intended to automate/stre the weekly reporting requirements.
"""


import argparse
import csv
import datetime
import json
import os
import quip


def get_quip_environment(default_settings_file='config_settings.json'):
    """Get commandline args and/or parse config file for intended operation.

    usage: quip_puller.py [-h] [-v] [--quip_api_key QUIP_API_KEY]
                          [--quip_doc_id QUIP_DOC_ID]
                          [config_file]

    Convert a quip document to misc report formats.

    positional arguments:
    config_file                   The configuration file specs.

    optional arguments:
    -h, --help                    show this help message and exit
    -v, --verbose
    --quip_api_key QUIP_API_KEY   API key for quip.
    --quip_doc_id QUIP_DOC_ID     Quip document ID.
    """
    parser = argparse.ArgumentParser(description='Convert a quip document '
                                     'to misc report formats.')
    parser.add_argument('-v', '--verbose', action='store_true')
    parser.add_argument('config_file', type=str, nargs='?',
                        default=default_settings_file,
                        help='The configuration file specs.')
    parser.add_argument('--quip_api_key', type=str, action='store',
                        help='API key for quip.')
    parser.add_argument('--quip_doc_id', type=str, action='store',
                        help='Quip document ID.')
    args = parser.parse_args()
    try:
        with open(args.config_file) as f:
            settings = json.loads(f.read())
    except (IOError):
        print('{} does not exsist.'.format(args.config_file))
        exit()
    # Default behaviour is to override if api or quip id explicitly passed.
    if args.quip_api_key is not None:
        settings['quip_api_key'] = args.quip_api_key
    if args.quip_doc_id is not None:
        settings['quip_doc_id'] = args.quip_doc_id

    print('Trying:\n  doc:{}\n  key:{}\n'.format(settings['quip_doc_id'],
                                                 settings['quip_api_key']))
    return (settings['quip_api_key'], settings['quip_doc_id'], args.verbose)


def dictReader(access_token, thread_id):
    """Import and parse a colaberative quip document.

    Specifically parse any spreadsheet and return a datastructure similar to
    that returned by csv.dictReader. Read more at:
    https://fb.quip.com/api/reference#threads-edit-document
    """
    def normalize_field(name):
        """Column names have inconsistent spaces and caps, so normalize."""
        return name.replace(' ', '_').lower()

    client = quip.QuipClient(access_token=access_token)
    print('Loading {}'.format(client.get_thread(thread_id)['thread']['title']))
    spreadsheet_tree = client.get_first_spreadsheet(thread_id)
    spreadsheet = client.parse_spreadsheet_contents(spreadsheet_tree)
    rows = spreadsheet['rows']
    results = []
    for row in rows[1:]:
        results.append({c: d['content'] for c, d in row['cells'].iteritems()})

    column_names = {column_id: normalize_field(name['content'])
                    for column_id, name in rows[0]['cells'].iteritems()}

    return [dict(zip(map(lambda x: column_names[x], r.keys()), r.values()))
            for r in results]


def generate_wiki_table(projects):
    """Built a string formated as a wiki table."""
    table_header = ('{| class="wikitable"\n'
                    '! <b>Project Name (link)</b>\n'
                    '! <b>OPM</b>\n'
                    '! <b>Overall Health</b>\n'
                    '! <span title="Meeting Deadlines?"><b>Time</b></span>\n'
                    '! <span title="Scope Alignment"><b>Scope</b></span>\n'
                    '! <span title="Potential Risks"><b>Risk</b></span>\n'
                    '! <b>Background Detail</b>\n'
                    '! <span title="Weeks Updates"><b>Updates</b></span>\n'
                    '! <b>Last Updated</b>\n\n')

    table_row = ('|-\n'
                 '| [[{url} {title}]]\n'
                 '| {project_manager}\n'
                 '| {overall_status}\n'
                 '| {time_status}\n'
                 '| {scope_status}\n'
                 '| {risk_status}\n'
                 '|\n'
                 '<div class="toccolours mw-collapsible mw-collapsed">\n'
                 '<br> <b>{priority}</b> priority <br> Click for details.'
                 '<div class="mw-collapsible-content">\n'
                 'It is sponsored by {sponsor}. It will be completed '
                 '<b>{target_completion}</b> based on a time allocation of '
                 '<b>{time_required}</b> this quarter.\n'
                 '----\n'
                 '<nowiki>\n{objectives}\n</nowiki>\n'
                 '</div>\n'
                 '</div>\n'
                 '|\n'
                 '<nowiki>\n{updates}\n</nowiki>\n'
                 '| {date_updated}\n\n')

    table_end = '|}'

    def wiki_status(field_value):
        """Convert a status into an appropriate image for the wiki."""
        image_map = {'r': '[[Image:RedCircle.png|20px]]',
                     'y': '[[Image:YellowCircle.png|20px]]',
                     'g': '[[Image:GreenCircle.png|20px]]',
                     'blocked': '[[Image:RedCircle.png|20px]]',
                     'on hold': '[[Image:YellowCircle.png|20px]]',
                     'in progress': '[[Image:GreenCircle.png|20px]]',
                     'in motion': '[[Image:GreenCircle.png|20px]]',
                     'complete': '[[Image:GreenCircle.png|20px]]'}
        try:
            return image_map[field_value.lower()]
        except(KeyError):
            return 'NA'

    def wiki_name(field_value):
        """Return a wiki formated employee reference."""
        unix_name = {'tad': 'thickman',
                     'chad': 'cshields',
                     'manish': 'manishk',
                     'kevin': 'kjfederation',
                     'geir': 'ghogberg',
                     'pritesh': 'pritesh',
                     'gene': 'genepage',
                     'garland': 'gje',
                     'jocy': 'jocy',
                     'paul': 'paulturner',
                     'david': 'davidt',
                     'jesse': 'jessepeters',
                     'ewa': 'suliga',
                     'john': 'jnovella',
                     'deb': 'debottym',
                     'nhan': 'nhantrieu',
                     'siv': 'siv',
                     'anusha': 'aalishetty',
                     'ciaran': 'thickman',  # Placeholder until Unix name avail
                     'lulu': 'thickman',  # Placeholder until Unix name avail
                     'None': 'unknown'}
        return '{{#person:' + unix_name.get(field_value.lower(), 'None') + '}}'

    table_rows = ''
    for project in projects:
        table_rows += table_row.format(
            url=project.get('url', 'https://fburl.com/ens_project_status'),
            title=project.get('title', 'None'),
            project_manager=wiki_name(project.get('project_manager', 'None')),
            overall_status=wiki_status(project.get('overall_status', 'g')),
            time_status=wiki_status(project.get('time_status', 'g')),
            scope_status=wiki_status(project.get('scope_status', 'g')),
            risk_status=wiki_status(project.get('risk_status', 'g')),
            priority=project.get('priority', 'None'),
            sponsor=wiki_name(project.get('sponsor', 'None')),
            target_completion=project.get('target_completion', 'None'),
            time_required=project.get('time_required', 'None'),
            objectives=project.get('objectives', 'None'),
            updates=project.get('updates', 'None'),
            date_updated=project.get('date_updated', datetime.date))

    return table_header + table_rows + table_end


def generate_markdown(projects):
    """Built a string formated as a markdown post.

    This function will perform all actions needed to create a string that can
    be uploaded, saved, modified.
    """
    def normalize_syntax(s):
        """Update a string so it can be displayed in markdown."""
        s = s.replace('#*', '*')
        s = s.replace('#', '*')
        return s

    project_markdown = ('# [{title}]({url})\n'
                        '## Report Updated on {date_updated}\n'
                        '## ( {priority} Priority and PMed by '
                        '{project_manager})\n{updates}\n'
                        '* * *\n\n')
    post = ''
    for project in projects:
        post += project_markdown.format(
            url=project.get('url', 'https://fburl/ens_project_status'),
            title=project.get('title', 'Unknown'),
            date_updated=project.get('date_updated', datetime.date),
            priority=project.get('priority', 'Unknown'),
            project_manager=project.get('project_manager', 'Unknown'),
            updates=normalize_syntax(project.get('updates', 'Unknown')))

    return post


def simple_display(projects):
    """Generate a pretty format of the data structure."""
    s = ''
    for project in projects:
        for field_name, contents in project.iteritems():
            s += '{} => {:.50}\n'.format(field_name, contents)

        s += '-------------------\n\n'
    return s


def include_in_reports(project):
    """Helper function to return if a project is included in a report."""
    if project.get('overall_status') == 'Complete':
        return False
    if project.get('updates') == 'None':
        return False
    if (project.get('ignore_in_reports', 'f').lower() == 't') or \
       (project.get('ignore_in_reports', 'f').lower() == 'true'):
        return False

    return True


def save_files(projects, file_name=''):
    """Save as properly formated text files."""
    if not os.path.exists('output'):
        os.makedirs('output')
    with open('output/{}markdown_export.txt'.format(file_name), 'w')as f:
        f.write(generate_markdown(projects))
    with open('output/{}wiki_export.txt'.format(file_name), 'w')as f:
        f.write(generate_wiki_table(projects))
    with open('output/{}mycsvfile.csv'.format(file_name), 'wb') as f:
        w = csv.DictWriter(f, projects[0].keys())
        w.writeheader()
        w.writerows(projects)


if __name__ == '__main__':
    api_key, doc_id, verbose = get_quip_environment('config_settings.json')
    projects = dictReader(access_token=api_key, thread_id=doc_id)
    filtered_projects = filter(include_in_reports, projects)

    print('{projects} projects in the quip doc.\n{filtered_projects} '
          'after filtering.'.format(projects=len(projects),
                                    filtered_projects=len(filtered_projects)))

    if verbose:
        print('-------------------\nWIKI FORMAT     \n-------------------\n\n'
              '{wiki}\n\n'
              '-------------------\nMARKDOWN FORMAT \n-------------------\n\n'
              '{markdown}\n\n'
              '-------------------\nSIMPLE FORMAT   \n-------------------\n\n'
              '{simple}').format(wiki=generate_wiki_table(filtered_projects),
                                 markdown=generate_markdown(filtered_projects),
                                 simple=simple_display(filtered_projects))

    save_files(filtered_projects, 'filered_')
    save_files(projects, 'unfilered_')
    print('\n\nFiles saved to the output folder.\n')

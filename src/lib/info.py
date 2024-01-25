from lib.config_manager import read_config_from_text_box


def group_consecutive_numbers(numbers):
    ranges = []
    for n in numbers:
        if not ranges or n > ranges[-1][-1] + 1:
            ranges.append([n])
        else:
            ranges[-1].append(n)
    return ['{}-{}'.format(r[0], r[-1]) if len(r) > 1 else str(r[0]) for r in ranges]


def get_all_senders(prs):
    all_senders = []
    for i, slide in enumerate(prs.slides):
        # image_name = read_config(slide, "image_name")
        sender = read_config_from_text_box(slide, "sender")
        all_senders.append(sender)
    return all_senders


def get_sender_positions(all_senders):
    senders = [x for i, x in enumerate(
        all_senders) if i == 0 or x != all_senders[i-1]]
    duplicates = [item for item in senders if senders.count(item) > 1]
    senders = list(dict.fromkeys(duplicates))

    sender_str_list = []
    for sender in senders:
        positions = [i+1 for i, x in enumerate(all_senders) if x == sender]
        grouped_positions = group_consecutive_numbers(positions)
        sender_str_list.append(
            f'{sender} appears at pages {", ".join(grouped_positions)}')
    return sender_str_list


def print_duplicate_senders(prs):
    all_senders = get_all_senders(prs)
    sender_positions = get_sender_positions(all_senders)
    if sender_positions:
        print("The following senders appear in more than one section in the presentation:")
        print("\n".join(sender_positions))


def print_new_pages_start(page_count, starting_page_count):
    if page_count > starting_page_count and starting_page_count > 0:
        print(f'New pages start from page: {starting_page_count + 1}')

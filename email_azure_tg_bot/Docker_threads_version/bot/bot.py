from telebot import TeleBot, types
from imapclient import IMAPClient
import email
from email import utils
from email.parser import Parser
from email.policy import default
import threading
from time import sleep
from datetime import datetime, timedelta
from reg_item import register_problem
from email_worker import send_email, parse_email_body
import os
import configparser


# Read configuration files
config = configparser.ConfigParser()
config.read_file(open(r'bot.cfg', encoding='utf-8'))
imap_host = config.get('bot', 'imap_host')
imap_port = int(config.get('bot', 'imap_port'))
imap_username = config.get('bot', 'imap_username')
imap_password = config.get('bot', 'imap_password')
t_token = config.get('bot', 't_token')
allowed_users = [int(i) for i in config.get('bot', 'allowed_users').split(',')]
admin = allowed_users[0]
project_names = dict()
for i in config.get('bot', 'project_names').split(';'):
    project_names[i.split(':')[0]] = tuple(i.split(':')[1].split(','))
bot = TeleBot(t_token)


# Reset/clear information for email registration
def clear_reg_email_pool(user_id):
    del check_dict[int(user_id)][1][:]
    del check_dict[int(user_id)][2]['sent_to'][:]
    check_dict[int(user_id)][2]['reg_flag'] = True
    check_dict[int(user_id)][2]['send_flag'] = True
    check_dict[int(user_id)][2]['reg_item_num'] = True


# Check_mail function (connect, parse, reply)
def read_email(tg_id):
    check_dict[tg_id][3]['process_started'] = datetime.now()
    bot.send_message(tg_id, 'Email check started.')
    while process_dict[tg_id][0] == True:
        try:
            # Connection
            client = IMAPClient(imap_host, port=imap_port, use_uid=True, ssl=True)
            client.login(imap_username, imap_password)
            client.select_folder('INBOX', readonly=True)
            emessages = client.search('UNSEEN')[-30:]
            response = dict(client.fetch(emessages, ['RFC822']))

            # Clear read_massages information
            if len(check_dict[tg_id][0]) > 500:
                check_dict[tg_id][0][:] = check_dict[tg_id][0][50:]

            # Parsing
            for i in range(len(emessages)):
                emessage_str = response[emessages[i]][b'RFC822'].decode("utf-8")
                msg_dict = Parser(policy=default).parsestr(emessage_str)
                msg_date_time = '{:%A, %d.%m.%Y %H:%M}'.format(email.utils.parsedate_to_datetime(msg_dict['Date'])
                                                               + timedelta(hours=(3 - int(msg_dict['Date'][-4:-2]))))
                msg_time = '{:%H:%M:%S}'.format(email.utils.parsedate_to_datetime(msg_dict['Date'])
                                                + timedelta(hours=(3 - int(msg_dict['Date'][-4:-2]))))
                msg_date = '{:%d.%m.%Y}'.format(email.utils.parsedate_to_datetime(msg_dict['Date'])
                                                + timedelta(hours=(3 - int(msg_dict['Date'][-4:-2]))))

                if msg_dict['Message-ID'] not in check_dict[tg_id][0]:
                    check_dict[tg_id][3]['read_emails'] += 1
                    bot_msg = ('From: ' + msg_dict['From'] + '\n' + 'To: ' + msg_dict['To'] + '\n' +
                               (('СС: ' + msg_dict['СС'] + '\n') if msg_dict['СС'] else '') + 'Date: ' +
                               msg_date_time + '\n' + 'Subject: ' + msg_dict['Subject'] + '\n')
                    email_body = email.message_from_string(emessage_str)
                    body = parse_email_body(email_body)
                    bot_msg = bot_msg + '\n' + body + ' ...'

                    # Replying
                    try:
                        markup = types.InlineKeyboardMarkup()
                        check_dict[tg_id][1].append({'msg_num': i,
                                                     'msg_time': msg_time,
                                                     'msg_date': msg_date,
                                                     'bot_msg': bot_msg,
                                                     'msg_title': msg_dict['Subject'][:50]})

                        btn1 = types.InlineKeyboardButton('Register',
                                                          callback_data='start_register' + '|system|'
                                                                        + str(tg_id) + '|' + str(i))
                        markup.add(btn1)
                        bot.send_message(tg_id, bot_msg, reply_markup=markup)
                        check_dict[tg_id][0].append(msg_dict['Message-ID'])
                        check_dict[tg_id][2]['sent_to'].append(str(msg_dict['To']))
                    except Exception as exc:
                        print(exc)
                        print('{} / Tg_ID: {} - Exception in thread: bot send message.'.format(datetime.now(), tg_id))
                        check_dict[tg_id][3]['errors'] += 1
                        continue

            client.logout()
            for _ in range(120):
                if not process_dict[tg_id][0]:
                    break
                else:
                    sleep(0.5)

        except Exception as exc:
            print(exc)
            print('{} / Tg_ID: {} - Exception in thread: while cycle'.format(datetime.now(), tg_id))
            check_dict[tg_id][3]['errors'] += 1
            continue


# Callback handler for registration and send reply messages
@bot.callback_query_handler(func=lambda call: True)
def callback_inline(call):
    user_id = call.data.split('|')[2]
    message_id = call.data.split('|')[3]
    um_ids = '|' + user_id + '|' + message_id

    # Start registration
    if (call.data.startswith('start_register') and check_dict[int(user_id)][1]
            and check_dict[int(user_id)][2]['reg_flag']):
        check_dict[int(user_id)][2]['reg_flag'] = False
        markup = types.InlineKeyboardMarkup()
        btn1 = types.InlineKeyboardButton('Project_1', callback_data='register_0|proj1' + um_ids)
        btn2 = types.InlineKeyboardButton('Project_2', callback_data='register_0|proj2' + um_ids)
        btn3 = types.InlineKeyboardButton('Project_3', callback_data='register_0|proj3' + um_ids)
        btn4 = types.InlineKeyboardButton('Project_4', callback_data='register_0|proj4' + um_ids)
        btn5 = types.InlineKeyboardButton('Project_5', callback_data='register_0|proj5' + um_ids)
        btn6 = types.InlineKeyboardButton('Abort', callback_data='stop|any' + um_ids)
        markup.add(btn1, btn2, btn3, btn4, btn5, btn6)
        bot.send_message(call.message.chat.id, 'Select project to register.', reply_markup=markup)

    # Registration_0
    elif call.data.startswith('register_0') and check_dict[int(user_id)][2]['send_flag']:
        project = call.data.split('|')[1]
        if project_names[project][1] in check_dict[int(user_id)][2]['sent_to'][int(message_id)]:
            markup = types.InlineKeyboardMarkup()
            btn1 = types.InlineKeyboardButton('Yes', callback_data='register_1|' + project + um_ids)
            btn2 = types.InlineKeyboardButton('No', callback_data='stop|any' + um_ids)
            markup.add(btn1, btn2)
            bot.send_message(call.message.chat.id,
                             'Register request in the project {}?'.format(project_names[project][0]),
                             reply_markup=markup)
        else:
            markup = types.InlineKeyboardMarkup()
            btn1 = types.InlineKeyboardButton('Yes', callback_data='register_1|' + project + um_ids)
            btn2 = types.InlineKeyboardButton('No', callback_data='stop|any' + um_ids)
            markup.add(btn1, btn2)
            bot.send_message(call.message.chat.id, '!ATTENTION! The address indicated in the letter does not correspond '
                                                   'to the draft registration. Are you sure you want to register a '
                                                   'request in the project {}?'.format(project_names[project][0]), reply_markup=markup)

    # Registration_1
    elif call.data.startswith('register_1') and check_dict[int(user_id)][2]['send_flag']:
        check_dict[int(user_id)][2]['send_flag'] = False
        project = call.data.split('|')[1]
        try:
            reg_item_num = register_problem(project=project,
                                            time=check_dict[int(user_id)][1][int(message_id)]['msg_time'],
                                            title=check_dict[int(user_id)][1][int(message_id)]['msg_title'],
                                            text='<div>' + check_dict[int(user_id)][1][int(message_id)][
                                                'bot_msg'] + '</div>')
            check_dict[int(user_id)][2]['reg_item_num'] = reg_item_num
            check_dict[int(user_id)][3]['reg_emails'] += 1
            markup = types.InlineKeyboardMarkup()
            btn1 = types.InlineKeyboardButton('Yes', callback_data='send|' + project + um_ids)
            btn2 = types.InlineKeyboardButton('No', callback_data='stop|' + project + um_ids)
            markup.add(btn1, btn2)
            bot.send_message(call.message.chat.id,
                             'The problem is registered in project {} with number - {}. Send a corresponding letter '
                             'about registration?'.format(project_names[project][0], reg_item_num), reply_markup=markup)
        except Exception as exc:
            clear_reg_email_pool(user_id)
            print(exc)
            print(str(call.data) + 'Registration failed!')
            check_dict[int(user_id)][3]['errors'] += 1
            bot.send_message(call.message.chat.id, 'Registration failed!')

    # End of registration
    elif call.data.startswith('send') and check_dict[int(user_id)][2]['reg_item_num']:
        project = call.data.split('|')[1]
        try:
            send_email(project=project,
                       reply_msg=check_dict[int(user_id)][1][int(message_id)]['bot_msg'],
                       reg_item_num=check_dict[int(user_id)][2]['reg_item_num'],
                       msg_date=check_dict[int(user_id)][1][int(message_id)]['msg_date'])
            clear_reg_email_pool(user_id)
            check_dict[int(user_id)][3]['sent_emails'] += 1
            bot.send_message(call.message.chat.id,
                             '{}. Registration email sent successfully.'.format(project_names[project][0]))
        except Exception as exc:
            clear_reg_email_pool(user_id)
            print(exc)
            print(str(call.data) + 'Sending failed!')
            check_dict[int(user_id)][3]['errors'] += 1
            bot.send_message(call.message.chat.id, 'Sending failed!')

    # Stop registration or sending
    elif call.data.startswith('stop'):
        if not check_dict[int(user_id)][2]['reg_flag']:
            clear_reg_email_pool(user_id)
            bot.send_message(call.message.chat.id, 'Aborted!')


# Start bot commands
@bot.message_handler(commands=['start'])
def start_bot(message):
    markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
    btn1 = types.KeyboardButton('Start')
    btn2 = types.KeyboardButton('Stop')
    btn3 = types.KeyboardButton('Restart')
    markup.row(btn1, btn2, btn3)
    bot.send_message(message.chat.id, 'Bot welcomes you!', reply_markup=markup)


# Statistics
@bot.message_handler(commands=['stats'])
def stats(message):
    if message.from_user.id == admin:
        total_errors = 0
        total_read = 0
        total_reg = 0
        total_send = 0
        for k, v in check_dict.items():
            total_errors += v[3]['errors']
            total_read += v[3]['read_emails']
            total_reg += v[3]['reg_emails']
            total_send += v[3]['sent_emails']
        state_msg = ('Bot started: {} \nErrors: {} \nTotal emails: {} \nRegistered: {} \nSent: {} \nCurrent '
                     'state: \n').format(bot_start, total_errors, total_read, total_reg, total_send)
        for k, v in check_dict.items():
            state_msg = (state_msg + '...{} - {}. Started/stopped: {}, Read: {}. Registered: {}. '
                                     'Sent: {}.'.format(str(k)[-3:], str(process_dict[k][0]),
                                                              str(process_dict[k][1]), str(v[3]['read_emails']),
                                                              str(v[3]['reg_emails']), str(v[3]['sent_emails'])) + '\n')
        bot.send_message(message.chat.id, state_msg)
    else:
        bot.reply_to(message, 'Access denied.')


# Stop bot
@bot.message_handler(commands=['finish'])
def finish(message):
    if message.from_user.id == admin:
        print('Bot stopped')
        for k, v in process_dict.items():
            if v[0]:
                process_dict[k][0] = False
                bot.send_message(k, 'Bot stopped by the administrator!. Email is not checking.')
        os._exit(0)
    else:
        bot.reply_to(message, 'Access denied.')


# Start email checking (threading for user)
@bot.message_handler(commands=['start_check'])
def start_check(message):
    if message.from_user.id in allowed_users:
        if not process_dict[message.from_user.id][0]:
            new_thread = threading.Thread(target=read_email, args=(message.from_user.id,), name='Parsing Thread')
            new_thread.start()
            process_dict[message.from_user.id] = [True, str(datetime.now())]
        else:
            bot.send_message(message.chat.id, 'Bot is already running')
    else:
        bot.reply_to(message, 'Access denied.')
        print('Attempt to start by unauthorized user with UID {}'.format(message.from_user.id))


# Stop email checking
@bot.message_handler(commands=['stop'])
def stop_check(message):
    if message.from_user.id in allowed_users and process_dict[message.from_user.id][0]:
        process_dict[message.from_user.id] = [False, str(datetime.now())]
        bot.send_message(message.chat.id, "Mail checking stopped.")
    else:
        bot.reply_to(message, 'Access denied.')


# Restart email checking (threading for user)
@bot.message_handler(commands=['restart'])
def restart_check(message):
    if message.from_user.id in allowed_users:
        del check_dict[message.from_user.id][0][:]
        clear_reg_email_pool(message.from_user.id)
        if process_dict[message.from_user.id][0]:
            process_dict[message.from_user.id] = [False, str(datetime.now())]
            sleep(0.5)
        bot.send_message(message.chat.id, "Restarting mail checking...")
    else:
        bot.reply_to(message, 'Access denied.')


# Button text handler
@bot.message_handler(content_types=['text'])
def btn_handler(message):
    if message.text == 'Start':
        try:
            start_check(message)
        except Exception as exc:
            print(exc)
            print('{} - Error in Check_mail function. Restarting...'.format(datetime.now()))
            check_dict[message.from_user.id][3]['errors'] += 1
    elif message.text == 'Stop':
        stop_check(message)
    elif message.text == 'Restart':
        restart_check(message)
        try:
            start_check(message)
        except Exception as exc:
            print(exc)
            print('{} - Error in Check_mail function. Restarting...'.format(datetime.now()))
            check_dict[message.from_user.id][3]['errors'] += 1
    else:
        print('Unrecognized command: ' + message.text)


if __name__ == '__main__':
    # Check email dictionary in format <<id: [send to tg user massages], [emails for registration], {registration control}, {stats}>>
    check_dict = dict()
    # User process dictionary in format <<id: [process, start/terminate time]>>
    process_dict = dict()
    for i in allowed_users:
        process_dict[i] = [False, '']
        check_dict[i] = list([list([]),
                                      list([]),
                                      dict({'reg_flag': True, 'send_flag': True,
                                                    'reg_item_num': False, 'sent_to': list([])}),
                                      dict({'process_started': False, 'errors': 0, 'read_emails': 0,
                                                    'reg_emails': 0, 'sent_emails': 0})])
    bot_start = datetime.now()

    # Start bot polling
    try:
        print('Bot started at: {}'.format(datetime.now()))
        bot.infinity_polling(timeout=10, long_polling_timeout=5, logger_level=None)
    except Exception as exc:
        print(exc)
        print('Bot infinity_polling error: {}'.format(datetime.now()))


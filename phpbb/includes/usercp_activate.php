<?php
/***************************************************************************
 *                            usercp_activate.php
 *                            -------------------
 *   begin                : Saturday, Feb 13, 2001
 *   copyright            : (C) 2001 The phpBB Group
 *   email                : support@phpbb.com
 *
 *   $Id: usercp_activate.php,v 1.6 2002/04/03 20:14:47 the_systech Exp $
 *
 *
 ***************************************************************************/

/***************************************************************************
 *
 *   This program is free software; you can redistribute it and/or modify
 *   it under the terms of the GNU General Public License as published by
 *   the Free Software Foundation; either version 2 of the License, or
 *   (at your option) any later version.
 *
 *
 ***************************************************************************/

if ( !defined('IN_PHPBB') )
{
	die('Hacking attempt');
	exit;
}

$sql = "SELECT user_id, user_email, user_newpasswd, user_lang  
	FROM " . USERS_TABLE . "
	WHERE user_actkey = '" . str_replace("\'", "''", $HTTP_GET_VARS['act_key']) . "'";
if ( $result = $db->sql_query($sql) )
{
	if ( $row = $db->sql_fetchrow($result) )
	{
		$sql_update_pass = ( $row['user_newpasswd'] != '' ) ? ", user_password = '" . str_replace("\'", "''", $row['user_newpasswd']) . "', user_newpasswd = ''" : "";

		$sql = "UPDATE " . USERS_TABLE . "
			SET user_active = 1, user_actkey = ''" . $sql_update_pass . " 
			WHERE user_id = " . $row['user_id'];
		if ( $result = $db->sql_query($sql) )
		{
			if ( $board_config['require_activation'] == USER_ACTIVATION_ADMIN && $sql_update_pass == '' )
			{
				include($phpbb_root_path . 'includes/emailer.'.$phpEx);
				$emailer = new emailer($board_config['smtp_delivery']);

				$email_headers = 'From: ' . $board_config['board_email'] . "\r\nReturn-Path: " . $board_config['board_email'] . "\r\n";

				$emailer->use_template('admin_welcome_activated', $row['user_lang']);
				$emailer->email_address($row['user_email']);
				$emailer->set_subject();//$lang['Account_activated_subject']
				$emailer->extra_headers($email_headers);

				$emailer->assign_vars(array(
					'SITENAME' => $board_config['sitename'], 
					'USERNAME' => $username,
					'PASSWORD' => $password_confirm,
					'EMAIL_SIG' => str_replace('<br />', "\n", "-- \n" . $board_config['board_email_sig']))
				);
				$emailer->send();
				$emailer->reset();

				$template->assign_vars(array(
					'META' => '<meta http-equiv="refresh" content="10;url=' . append_sid("index.$phpEx") . '">')
				);

				message_die(GENERAL_MESSAGE, $lang['Account_active_admin']);
			}
			else
			{
				$template->assign_vars(array(
					'META' => '<meta http-equiv="refresh" content="10;url=' . append_sid("index.$phpEx") . '">')
				);

				$message = ( $sql_update_pass == '' ) ? $lang['Account_active'] : $lang['Password_activated']; 
				message_die(GENERAL_MESSAGE, $message);
			}
		}
		else
		{
			message_die(GENERAL_ERROR, 'Could not update users table', '', __LINE__, __FILE__, $sql_update);
		}
	}
	else
	{
		message_die(GENERAL_ERROR, $lang['Wrong_activation']); //wrongactiv
	}
}
else
{
	message_die(GENERAL_ERROR, 'Could not obtain user information', '', __LINE__, __FILE__, $sql);
}

?>

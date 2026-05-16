        "ark_api_key": "ARK_API_KEY",
        "ark_api_base": "ARK_API_BASE",
        "dashscope_api_key": "DASHSCOPE_API_KEY",
        "dashscope_api_base": "DASHSCOPE_API_BASE",
        # Channel credentials (used by skills that check env vars)
        "feishu_app_id": "FEISHU_APP_ID",
        "feishu_app_secret": "FEISHU_APP_SECRET",
        "dingtalk_client_id": "DINGTALK_CLIENT_ID",
        "dingtalk_client_secret": "DINGTALK_CLIENT_SECRET",
        "wechatmp_app_id": "WECHATMP_APP_ID",
        "wechatmp_app_secret": "WECHATMP_APP_SECRET",
        "wechatcomapp_agent_id": "WECHATCOMAPP_AGENT_ID",
        "wechatcomapp_secret": "WECHATCOMAPP_SECRET",
        "qq_app_id": "QQ_APP_ID",
        "qq_app_secret": "QQ_APP_SECRET",
        "weixin_token": "WEIXIN_TOKEN",
    }
    injected = 0
    for conf_key, env_key in _CONFIG_TO_ENV.items():
        if env_key not in os.environ:
            val = config.get(conf_key, "")
            if val:
                os.environ[env_key] = str(val)
                injected += 1

    injected += _sync_skill_config_to_env(config.get("skill", {}))

    if injected:
        logger.info("[INIT] Synced {} config values to environment variables".format(injected))

    config.load_user_datas()


def _sync_skill_config_to_env(skill_section) -> int:
    """Flatten skill-namespaced config into environment variables.

    Mapping rule: ``config["skill"][<name>][<key>]`` -> ``SKILL_<NAME>_<KEY>``
    (e.g. ``skill["image-generation"].model`` -> ``SKILL_IMAGE_GENERATION_MODEL``).

    This lets subprocess-based skill scripts read their own settings without
    importing project code. Existing env vars are NOT overwritten so the
    real environment always wins.

    Returns the number of variables actually injected.
    """
    if not isinstance(skill_section, dict):
        return 0
    injected = 0
    for skill_name, skill_conf in skill_section.items():
        if not isinstance(skill_conf, dict):
            continue
        name_part = str(skill_name).replace("-", "_").upper()
        for key, val in skill_conf.items():
            if val is None or val == "":
                continue
            env_key = "SKILL_{}_{}".format(name_part, str(key).upper())
            if env_key in os.environ:
                continue
            os.environ[env_key] = str(val)
            injected += 1
    return injected


def get_root():
    return os.path.dirname(os.path.abspath(__file__))


def read_file(path):
    with open(path, mode="r", encoding="utf-8-sig") as f:
        return f.read()


def conf():
    return config


def get_appdata_dir():
    data_path = os.path.join(get_root(), conf().get("appdata_dir", ""))
    if not os.path.exists(data_path):
        logger.info("[INIT] data path not exists, create it: {}".format(data_path))
        os.makedirs(data_path)
    return data_path


def subscribe_msg():
    trigger_prefix = conf().get("single_chat_prefix", [""])[0]
    msg = conf().get("subscribe_msg", "")
    return msg.format(trigger_prefix=trigger_prefix)


# global plugin config
plugin_config = {}


def write_plugin_config(pconf: dict):
    """
    :param pconf: 全量插件配置
    """
    global plugin_config
    for k in pconf:
        plugin_config[k.lower()] = pconf[k]

def remove_plugin_config(name: str):
    """
    :param name: 待重载的插件名
    """
    global plugin_config
    plugin_config.pop(name.lower(), None)


def pconf(plugin_name: str) -> dict:
    """
    :param plugin_name: 插件名称
    :return: 该插件的配置项
    """
    return plugin_config.get(plugin_name.lower())


global_config = {"admin_users": []}

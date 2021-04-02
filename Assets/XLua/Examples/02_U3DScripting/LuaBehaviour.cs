/*
 * Tencent is pleased to support the open source community by making xLua available.
 * Copyright (C) 2016 THL A29 Limited, a Tencent company. All rights reserved.
 * Licensed under the MIT License (the "License"); you may not use this file except in compliance with the License. You may obtain a copy of the License at
 * http://opensource.org/licenses/MIT
 * Unless required by applicable law or agreed to in writing, software distributed under the License is distributed on an "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied. See the License for the specific language governing permissions and limitations under the License.
*/

using UnityEngine;
using XLua;
using System;

namespace XLuaTest
{
    [System.Serializable]
    public class Injection
    {
        //通过一个name-value组合，注射【显示对象】到绑定的Lua类中
        public string name;
        public GameObject value;
    }

    [LuaCallCSharp]
    public class LuaBehaviour : MonoBehaviour
    {
        //Lua脚本作为文本资源，绑定到【预制体】的【C#脚本】上
        public TextAsset luaScript;
        //【显示对象】，通过【Injection】数组绑定到【预制体】的【C#脚本】上
        public Injection[] injections;

        internal static LuaEnv luaEnv = new LuaEnv(); //all lua behaviour shared one luaenv only!
        internal static float lastGCTime = 0;
        internal const float GCInterval = 1;//1 second 

        private Action luaStart;
        private Action luaUpdate;
        private Action luaOnDestroy;

        private LuaTable scriptEnv;

        void Awake()
        {
            scriptEnv = luaEnv.NewTable();

            //为每个脚本设置一个独立的环境，可一定程度上防止脚本间全局变量、函数冲突
            LuaTable meta = luaEnv.NewTable();
            meta.Set("__index", luaEnv.Global);
            scriptEnv.SetMetaTable(meta);
            meta.Dispose();

            //1、首先，把当前【C#脚本】组件对象，映射到Lua中的全局变量self
            scriptEnv.Set("self", this);
            //2、把注射的显示对象，按照设置的name，映射到Lua中的全局变量
            foreach (var injection in injections){
                scriptEnv.Set(injection.name, injection.value);
            }

            luaEnv.DoString(luaScript.text, "LuaTestScript", scriptEnv);

            Action luaAwake = scriptEnv.Get<Action>("awake");
            scriptEnv.Get("start", out luaStart);
            scriptEnv.Get("update", out luaUpdate);
            scriptEnv.Get("ondestroy", out luaOnDestroy);

            if (luaAwake != null)
            {
                luaAwake();
            }
        }

        // Use this for initialization
        void Start()
        {
            if (luaStart != null)
            {
                luaStart();
            }
        }

        // Update is called once per frame
        void Update()
        {
            if (luaUpdate != null)
            {
                luaUpdate();
            }
            if (Time.time - LuaBehaviour.lastGCTime > GCInterval)
            {
                luaEnv.Tick();
                LuaBehaviour.lastGCTime = Time.time;
            }
        }

        void OnDestroy()
        {
            if (luaOnDestroy != null)
            {
                luaOnDestroy();
            }
            luaOnDestroy = null;
            luaUpdate = null;
            luaStart = null;
            scriptEnv.Dispose();
            injections = null;
        }
    }
}
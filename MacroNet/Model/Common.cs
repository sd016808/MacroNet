using System;

namespace MacroNet.Model
{
    public enum CommandType
    {
        Mouse,
        Keyboard,
        Delay
    }

    [Flags]
    public enum MouseEventFlags
    {
        LeftDown = 0x00000002,
        LeftUp = 0x00000004,
        MiddleDown = 0x00000020,
        MiddleUp = 0x00000040,
        Move = 0x00000001,
        Absolute = 0x00008000,
        RightDown = 0x00000008,
        RightUp = 0x00000010,
        LeftClick = 0x00000080,
        RightClick = 0x00000100,
    }

    public enum KeyBoardEvent
    {
        Key_Down,
        Key_Up
    }
}

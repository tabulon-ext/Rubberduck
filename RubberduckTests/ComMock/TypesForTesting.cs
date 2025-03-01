﻿using System;

namespace RubberduckTests.ComMock
{

    public interface ITest1
    {
        void Do();
        int DoThis();
    }

    public interface ITest2
    {
        string DoThat();
    }

    public interface ITest3
    {
        void DoInt(int something);
        void DoString(string something);
        void DoObject(object something);
    }

    public interface ITestRef
    {
        void DoInt(ref int something);
        void DoString(ref string something);
        void DoObject(ref object something);

        int ReturnInt(ref int something);
    }

    public class ConvertibleTest : IConvertible
    {
        private readonly TypeCode _code;

        public ConvertibleTest(TypeCode code)
        {
            _code = code;
        }

        public TypeCode GetTypeCode()
        {
            return _code;
        }

        public bool ToBoolean(IFormatProvider provider)
        {
            return true;
        }

        public char ToChar(IFormatProvider provider)
        {
            return 't';
        }

        public sbyte ToSByte(IFormatProvider provider)
        {
            return 1;
        }

        public byte ToByte(IFormatProvider provider)
        {
            return 1;
        }

        public short ToInt16(IFormatProvider provider)
        {
            return 1;
        }

        public ushort ToUInt16(IFormatProvider provider)
        {
            return 1;
        }

        public int ToInt32(IFormatProvider provider)
        {
            return 1;
        }

        public uint ToUInt32(IFormatProvider provider)
        {
            return 1;
        }

        public long ToInt64(IFormatProvider provider)
        {
            return 1;
        }

        public ulong ToUInt64(IFormatProvider provider)
        {
            return 1;
        }

        public float ToSingle(IFormatProvider provider)
        {
            return 1;
        }

        public double ToDouble(IFormatProvider provider)
        {
            return 1;
        }

        public decimal ToDecimal(IFormatProvider provider)
        {
            return 1;
        }

        public DateTime ToDateTime(IFormatProvider provider)
        {
            return DateTime.MinValue;
        }

        public string ToString(IFormatProvider provider)
        {
            return "true";
        }

        public object ToType(Type conversionType, IFormatProvider provider)
        {
            return this;
        }
    }
}

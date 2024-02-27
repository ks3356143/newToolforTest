#python处理floatIEE754格式转换
import struct

def IEEE754_16_to_float(x,byte_type):
    if byte_type == 32:
        return struct.unpack('>f',struct.pack('>I',int(x,16)))[0]
    if byte_type == 64:
        return struct.unpack('>d',struct.pack('>Q',int(x,16)))[0]
    
def IEEE754_float_to_16(x,byte_type):
    if byte_type == 32:
        return hex(struct.unpack('>I',struct.pack('>f',x))[0])[2:].upper()
    if byte_type == 64:
        return hex(struct.unpack('>Q',struct.pack('>d',x))[0])[2:].upper()

if __name__ == '__main__':
    ##测试程序##
    x = 'BE051EB8'
    y = -0.13
    print(IEEE754_16_to_float(x,32))
    print(IEEE754_float_to_16(y,32))

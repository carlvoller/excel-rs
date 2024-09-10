use std::io::Read;

use csv::{ByteRecord, Reader};

pub fn bytes_to_csv<V: Read>(bytes: V) -> Reader<V> {
    csv::ReaderBuilder::new().from_reader(bytes)
}

pub fn get_headers<V: Read>(reader: &mut Reader<V>) -> Option<&ByteRecord> {
    match reader.byte_headers() {
        Ok(record) => Some(record),
        Err(_) => None,
    }
}

pub fn get_next_record<V: Read>(reader: &mut Reader<V>) -> Option<ByteRecord> {
    let mut record = csv::ByteRecord::new();
    match reader.read_byte_record(&mut record) {
        Ok(status) => {
            if status {
                Some(record)
            } else {
                None
            }
        }
        Err(_) => None,
    }
}

// #[cfg(test)]
// mod tests {
//     use super::*;

//     #[test]
//     fn it_works() {
//         let result = add(2, 2);
//         assert_eq!(result, 4);
//     }
// }

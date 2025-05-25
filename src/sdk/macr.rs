// #[macro_export]
// macro_rules! ENUM {
//     (
//         $(#[$outer:meta])*
//         enum $name:ident {
//             $($variant:ident = $value:expr,)+
//         }
//     ) => {
//         $(#[$outer])*
//         #[repr(i32)]
//         #[derive(Debug, Copy, Clone, PartialEq, Eq, Hash)]
//         pub enum $name {
//             $($variant = $value,)+
//         }
//         impl $name {
//             pub fn from_i32(value: i32) -> Option<Self> {
//                 match value {
//                     $(x if x == $name::$variant as i32 => Some($name::$variant),)+
//                     _ => None,
//                 }
//             }

//             pub fn to_i32(&self) -> i32 {
//                 *self as i32
//             }
//         }
//     };
// }